import sys

sys.dont_write_bytecode = True

import shutil
import tempfile
from pathlib import Path
from time import perf_counter


def print_header(title):
    """
    Print a section header.
    """
    print(title)
    print("-" * len(title))


def format_elapsed_time(seconds):
    """
    Format elapsed seconds as HH:MM:SS.
    """
    total_seconds = int(round(seconds))
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def cleanup_python_cache(project_folder):
    """
    Remove all __pycache__ folders and .pyc files inside the project,
    excluding virtual environments and other non-project folders.

    This should run before importing project modules.
    """
    project_folder = Path(project_folder)

    excluded_folder_names = {
        ".venv",
        "venv",
        "env",
        ".git",
        ".mypy_cache",
        ".pytest_cache",
    }

    removed_files = 0
    removed_folders = 0
    failed_items = []

    def is_excluded_path(path):
        path_parts = set(path.parts)

        return any(
            excluded_folder_name in path_parts
            for excluded_folder_name in excluded_folder_names
        )

    # 1. Remove .pyc files outside excluded folders.
    for pyc_file in project_folder.rglob("*.pyc"):
        if not pyc_file.is_file():
            continue

        if is_excluded_path(pyc_file):
            continue

        try:
            pyc_file.unlink()
            removed_files += 1
        except OSError as error:
            failed_items.append(
                {
                    "path": pyc_file,
                    "error": str(error),
                }
            )

    # 2. Remove __pycache__ folders outside excluded folders.
    pycache_folders = [
        folder
        for folder in project_folder.rglob("__pycache__")
        if folder.is_dir() and not is_excluded_path(folder)
    ]

    pycache_folders = sorted(
        pycache_folders,
        key=lambda folder: len(folder.parts),
        reverse=True,
    )

    for pycache_folder in pycache_folders:
        try:
            shutil.rmtree(pycache_folder)
            removed_folders += 1
        except OSError as error:
            failed_items.append(
                {
                    "path": pycache_folder,
                    "error": str(error),
                }
            )

    return {
        "removed_files": removed_files,
        "removed_folders": removed_folders,
        "failed_items": failed_items,
    }


def print_cache_cleanup_result(cache_cleanup_result, title="Python cache cleanup"):
    """
    Print Python cache cleanup result.
    """
    print(
        f"{title}: "
        f"{cache_cleanup_result['removed_files']} .pyc file(s) removed, "
        f"{cache_cleanup_result['removed_folders']} __pycache__ folder(s) removed."
    )

    if cache_cleanup_result["failed_items"]:
        print(f"{title} warnings:")

        for failed_item in cache_cleanup_result["failed_items"]:
            print(f"  Could not remove: {failed_item['path']}")
            print(f"  Reason: {failed_item['error']}")


def load_dependencies():
    """
    Import project dependencies after cache cleanup.

    This avoids trying to delete __pycache__ folders while imported modules
    are still being initialized or locked by the current Python process.
    """
    from core.config_reader import read_active_configuration
    from modules.AR.ar_001 import run_ar_001
    from modules.AR.ar_002 import run_ar_002
    from modules.AR.ar_003 import run_ar_003
    from modules.AR.ar_004 import run_ar_004
    from modules.AR.ar_005 import run_ar_005
    from modules.AR.ar_006 import run_ar_006
    from modules.CD.cd_001 import run_cd_001
    from modules.CD.cd_002 import run_cd_002

    control_runners = {
        "AR_001": run_ar_001,
        "AR_002": run_ar_002,
        "AR_003": run_ar_003,
        "AR_004": run_ar_004,
        "AR_005": run_ar_005,
        "AR_006": run_ar_006,
        "CD_001": run_cd_001,
        "CD_002": run_cd_002,

    }

    return read_active_configuration, control_runners


def create_temp_copy(file_path):
    """
    Create a temporary copy of an input workbook and return its path.

    This is useful for config.xlsx because the original workbook may be open
    in Excel. Reading from a temporary copy reduces lock-related issues.
    """
    file_path = Path(file_path)

    temp_folder = Path(tempfile.gettempdir()) / "lbr_analysis"
    temp_folder.mkdir(parents=True, exist_ok=True)

    temp_file = temp_folder / file_path.name

    shutil.copy2(file_path, temp_file)

    return temp_file


def cleanup_temp_config_copy(temp_config_path, original_config_path):
    """
    Delete temporary config copy when it is different from the original file.
    """
    if temp_config_path is None:
        return

    temp_config_path = Path(temp_config_path)
    original_config_path = Path(original_config_path)

    try:
        if temp_config_path.exists() and temp_config_path.resolve() != original_config_path.resolve():
            temp_config_path.unlink()
    except OSError:
        pass


def is_file_locked(file_path):
    """
    Return True if a file appears to be locked by another process.

    This is mainly useful on Windows when Excel has a workbook open.
    """
    file_path = Path(file_path)

    if not file_path.exists():
        return False

    try:
        with file_path.open("a+b"):
            return False
    except OSError:
        return True


def print_project_paths(project_folder, config_path, input_folder, output_folder):
    """
    Print main project paths.
    """
    print()
    print_header("LBR Automation Runner")
    print(f"Project folder: {project_folder}")
    print(f"Config file: {config_path}")
    print(f"Input folder: {input_folder}")
    print(f"Output folder: {output_folder}")
    print()


def validate_project_paths(config_path, input_folder, output_folder):
    """
    Validate required project paths before processing.
    """
    if not config_path.exists():
        print("ERROR: config.xlsx was not found.")
        print(f"Expected file: {config_path}")
        print()
        return False

    if not input_folder.exists():
        print("ERROR: input folder was not found.")
        print(f"Expected folder: {input_folder}")
        print()
        return False

    output_folder.mkdir(parents=True, exist_ok=True)

    return True


def print_active_modules(active_modules):
    """
    Print active modules from config.xlsx.
    """
    print_header("Active modules")

    if len(active_modules) == 0:
        print("No active modules found.")
        print()
        return

    for module in active_modules:
        print(
            f"{module['module']} | "
            f"From: {module['from']} | "
            f"To: {module['to']} | "
            f"Companies: {module['companies']}"
        )

    print()


def print_active_controls(active_controls):
    """
    Print active controls from config.xlsx.
    """
    print_header("Active controls")

    if len(active_controls) == 0:
        print("No active controls found.")
        print()
        return

    for control in active_controls:
        print(f"{control['module']} | {control['id_control']} | {control['name']}")

    print()


def run_active_controls(
    active_controls,
    modules_by_name,
    project_folder,
    config_path,
    input_folder,
    output_folder,
    control_runners,
):
    """
    Execute all active controls found in config.xlsx.
    """
    print_header("Execution")

    for control in active_controls:
        control_id = control["id_control"]

        if control_id not in control_runners:
            print(f"SKIPPED: No runner implemented for {control_id}")
            continue

        module_name = control["module"]

        context = {
            "project_folder": project_folder,
            "config_path": config_path,
            "input_folder": input_folder,
            "output_folder": output_folder,
            "module": modules_by_name[module_name],
            "control": control,
        }

        print(f"Running {control_id} - {control['name']}")

        control_start_time = perf_counter()

        try:
            control_runners[control_id](context)
        except PermissionError as error:
            print()
            print(f"ERROR while running {control_id}.")
            print("A file may be open in Excel or locked by another process.")
            print(f"Details: {error}")
            print()
            raise

        control_elapsed_time = perf_counter() - control_start_time

        print(
            f"{control_id} elapsed time: "
            f"{format_elapsed_time(control_elapsed_time)} "
            f"({control_elapsed_time:.2f} seconds)"
        )
        print()


def main():
    """
    Main automation runner.

    This script:
    1. Cleans old Python cache before importing project modules.
    2. Imports project dependencies after cleanup.
    3. Reads config.xlsx from a temporary copy.
    4. Detects active modules.
    5. Detects active controls.
    6. Executes registered Python runners.
    7. Prints elapsed time by control and total run.

    Note:
    - sys.dont_write_bytecode = True prevents new .pyc files from being written.
    - .venv is excluded from cache cleanup.
    """
    run_start_time = perf_counter()
    temp_config_path = None

    project_folder = Path(__file__).parent
    config_path = project_folder / "config.xlsx"
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    startup_cache_cleanup_result = cleanup_python_cache(project_folder)

    read_active_configuration, control_runners = load_dependencies()

    print_project_paths(
        project_folder=project_folder,
        config_path=config_path,
        input_folder=input_folder,
        output_folder=output_folder,
    )

    print_cache_cleanup_result(
        startup_cache_cleanup_result,
        title="Startup Python cache cleanup",
    )
    print()

    if not validate_project_paths(
        config_path=config_path,
        input_folder=input_folder,
        output_folder=output_folder,
    ):
        return

    try:
        try:
            temp_config_path = create_temp_copy(config_path)
        except OSError as error:
            print("ERROR: config.xlsx could not be copied for reading.")
            print("Please close it if it is open in Excel and try again.")
            print(f"Config file: {config_path}")
            print(f"Details: {error}")
            print()
            return

        active_configuration = read_active_configuration(temp_config_path)

        active_modules = active_configuration["active_modules"]
        active_controls = active_configuration["active_controls"]
        modules_by_name = active_configuration["modules_by_name"]

        print_active_modules(active_modules)

        if len(active_modules) == 0:
            return

        print_active_controls(active_controls)

        if len(active_controls) == 0:
            return

        run_active_controls(
            active_controls=active_controls,
            modules_by_name=modules_by_name,
            project_folder=project_folder,
            config_path=config_path,
            input_folder=input_folder,
            output_folder=output_folder,
            control_runners=control_runners,
        )

    finally:
        cleanup_temp_config_copy(
            temp_config_path=temp_config_path,
            original_config_path=config_path,
        )

        run_elapsed_time = perf_counter() - run_start_time

        print()
        print_header("Runner status")
        print("Finished.")
        print(
            f"Total elapsed time: "
            f"{format_elapsed_time(run_elapsed_time)} "
            f"({run_elapsed_time:.2f} seconds)"
        )
        print()


if __name__ == "__main__":
    main()
