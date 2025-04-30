import logging

def setup_logging():
    """
    Set up logging configuration for the application.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("attrition_report.log"),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger("attrition_logger")
    return logger