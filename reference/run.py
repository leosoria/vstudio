"""
Main automation runner.

This script is the entry point for the automation.

What it does:
1. Reads Config.xlsx.
2. Detects active modules from the CONFIG sheet.
3. Detects active controls from each active module sheet.
4. Executes the corresponding Python runner for each active control.

Implemented AR controls:
- AR_001
- AR_002
- AR_003

"""

from pathlib import Path

from core.config_reader import read_active_configuration
from modules.AR.ar_001 import run_ar_001
from modules.AR.ar_002 import run_ar_002
from modules.AR.ar_003 import run_ar_003


CONTROL_RUNNERS = {
    "AR_001": run_ar_001,
    "AR_002": run_ar_002,
    "AR_003": run_ar_003,
}


def print_header(title):
    """
    Print a section header.
    """
    print(title)
    print("-" * len(title))


def main():
    project_folder = Path(__file__).parent
    config_path = project_folder / "config.xlsx"
    input_folder = project_folder / "input"
    output_folder = project_folder / "output"

    print()
    print_header("LBR Automation Runner")
    print(f"Project folder: {project_folder}")
    print(f"Config file: {config_path}")
    print(f"Input folder: {input_folder}")
    print(f"Output folder: {output_folder}")
    print()

    if not config_path.exists():
        print("ERROR: Config.xlsx was not found.")
        print(f"Expected file: {config_path}")
        print()
        return

    if not input_folder.exists():
        print("ERROR: Input folder was not found.")
        print(f"Expected folder: {input_folder}")
        print()
        return

    output_folder.mkdir(parents=True, exist_ok=True)

    active_configuration = read_active_configuration(config_path)

    active_modules = active_configuration["active_modules"]
    active_controls = active_configuration["active_controls"]
    modules_by_name = active_configuration["modules_by_name"]

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

    print_header("Active controls")

    if len(active_controls) == 0:
        print("No active controls found.")
        print()
        return

    for control in active_controls:
        print(f"{control['module']} | {control['id_control']} | {control['name']}")

    print()

    print_header("Execution")

    for control in active_controls:
        control_id = control["id_control"]

        if control_id not in CONTROL_RUNNERS:
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
        CONTROL_RUNNERS[control_id](context)

    print()
    print_header("Runner status")
    print("Finished.")
    print()


if __name__ == "__main__":
    main()
