import logging
import os
from worktray_creation import create_worktray
from input_file_processment import validate_worktray
from google_forms_submission import submit_to_google_forms

# Configurable parameters
LOGS_DIRECTORY = "_logs"  # Directory to store logs
LOGS_FILE = os.path.join(LOGS_DIRECTORY, "process_logs.log")  # Log file path

def configure_logging():
    """
    Configures logging for the application.
    """
    if not os.path.exists(LOGS_DIRECTORY):
        os.makedirs(LOGS_DIRECTORY)

    logging.basicConfig(
        filename=LOGS_FILE,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

def main():
    """
    Main function to execute all modules in sequence:
    1. Create the worktray.
    2. Validate the worktray.
    3. Submit data to Google Forms.
    """
    try:
        # Configure logging
        configure_logging()
        logging.info("Starting the process.")

        # Step 1: Create the worktray
        logging.info("Step 1: Creating the worktray.")
        worktray_created = create_worktray()
        if worktray_created.empty:
            logging.error("Worktray creation failed. Process terminated.")
            return
        logging.info("Worktray created successfully.")

        # Step 2: Validate the worktray
        logging.info("Step 2: Validating the worktray.")
        validation_success = validate_worktray()
        if not validation_success:
            logging.error("Worktray validation failed. Process terminated.")
            return
        logging.info("Worktray validation completed successfully.")

        # Step 3: Submit data to Google Forms
        logging.info("Step 3: Submitting data to Google Forms.")
        submission_success = submit_to_google_forms()
        if not submission_success:
            logging.error("Google Forms submission failed. Process terminated.")
            return
        logging.info("Google Forms submission completed successfully.")

        # Final success message
        logging.info("All steps completed successfully.")

    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}", exc_info=True)

if __name__ == "__main__":
    main()