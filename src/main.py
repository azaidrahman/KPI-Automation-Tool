import sys
from GUI import create_gui
from test_runner import run_test

def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--test':
        run_test()
    else:
        create_gui()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        with open("../gui_error_log.txt", "a") as log_file:
            log_file.write(f"Error in application: {e}\n")
    finally:
        sys.exit(0)