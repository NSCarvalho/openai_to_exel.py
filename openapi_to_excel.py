import os
import subprocess
import sys

def check_python():
    try:
        subprocess.check_call([sys.executable, '--version'])
        print(f"Python is installed: {sys.executable}")
    except subprocess.CalledProcessError:
        print("Python is not installed. Please install Python from https://www.python.org/downloads/.")
        sys.exit(1)

def check_requirements():
    if not os.path.exists('requirements.txt'):
        print("The 'requirements.txt' file is missing. Please ensure it is in the same folder as the script.")
        sys.exit(1)
    print("Found 'requirements.txt'. Checking for dependencies...")

def create_virtual_environment():
    if not os.path.exists('env'):
        print("Creating virtual environment...")
        if sys.platform == "win32":
            subprocess.check_call([sys.executable, "-m", "venv", "env"])
        else:
            subprocess.check_call([sys.executable, "-m", "venv", "env"])
        print("Virtual environment created successfully.")
    else:
        print("Virtual environment already exists.")

def activate_virtual_environment():
    print("Activating virtual environment...")
    activate_script = os.path.join('env', 'Scripts', 'activate.bat') if sys.platform == "win32" else os.path.join('env', 'bin', 'activate')
    
    if not os.path.exists(activate_script):
        print(f"Could not find the activate script at {activate_script}.")
        sys.exit(1)

    # Running activation through a new command prompt to ensure environment is activated
    subprocess.check_call([activate_script, '&&', 'python', '--version'], shell=True)

def are_dependencies_installed():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "show", "openpyxl"])
        subprocess.check_call([sys.executable, "-m", "pip", "show", "requests"])
        print("Dependencies are already installed.")
        return True
    except subprocess.CalledProcessError:
        print("Dependencies are not installed.")
        return False

def install_dependencies():
    print("Installing required dependencies...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    print("Dependencies installed successfully.")

def check_openapi_file():
    if not os.path.exists('openapi.json'):
        print("The 'openapi.json' file is missing. Please place your OpenAPI file in this folder.")
        sys.exit(1)
    print("Found 'openapi.json' file. Ready to process it.")

def run_script():
    print("Running the OpenAPI to Excel conversion script...")
    subprocess.check_call([sys.executable, "openapi_to_excel_app.py"])
    print("The conversion is complete! Your Excel file has been generated.")

def show_help():
    print("\n### OpenAPI to Excel Script Help ###")
    print("This script will convert your OpenAPI specification file into an Excel file.")
    print("1. Ensure that you have Python installed. You can download it from https://www.python.org/downloads/.")
    print("2. Place the 'openapi.json' file (your OpenAPI specification) in the same folder as the script.")
    print("3. Make sure the 'requirements.txt' file is present in the folder.")
    print("4. Run the script by executing the following command:")
    print("   python openapi_to_excel.py")
    print("5. Follow the on-screen instructions to install dependencies, create and activate the virtual environment, and run the conversion.")

def main():
    print("Welcome to the OpenAPI to Excel Script Setup Guide!")

    # Check for Python installation
    check_python()

    # Check for requirements.txt
    check_requirements()

    # Create and activate the virtual environment
    create_virtual_environment()

    # Activate virtual environment
    activate_virtual_environment()

    # Install dependencies only if they are not already installed
    if not are_dependencies_installed():
        install_dependencies()

    # Check for OpenAPI file
    check_openapi_file()

    # Run the script
    run_script()

    # If the user needs help or assistance
    # show_help()

if __name__ == "__main__":
    main()
