import imaplib
import os
import email
import json
import configparser
import logging
import argparse

CREDENTIALS_FILE_PATH = "./credentials.json"

CONFIG_FILE_PATH = "./config.toml"

DRY_RUN = False

class InvalidCredentialsException(Exception):
    """
    Exception raised for errors in the credentials file.

    Args:
        fields (array): credential fields with error
        message (string): explanation of the error
    """

    def __init__(self, fields, message="Malformed credentials file found"):
        self.fields = fields
        self.message = message + ": {} not found".format(", ".join(fields))
        super().__init__(self.message)


class InvalidConfigException(Exception):
    """
    Exception raised for errors in the configuration file.

    Args:
        message (string): explanation of the error
    """

    def __init__(self, message="Malformed config file found"):
        self.message = message
        super().__init__(self.message)


def set_logger(config):
    logging.basicConfig()
    if config["log"]["level"] == 'INFO':
        logging.getLogger().setLevel(logging.INFO)
    if config["log"]["level"] == 'DEBUG':
        logging.getLogger().setLevel(logging.DEBUG)
    if config["log"]["level"] == 'WARNING':
        logging.getLogger().setLevel(logging.WARNING)
    if config["log"]["level"] == 'WARN':
        logging.getLogger().setLevel(logging.WARN)
    if config["log"]["level"] == 'ERROR':
        logging.getLogger().setLevel(logging.ERROR)
    if config["log"]["level"] == 'CRITICAL':
        logging.getLogger().setLevel(logging.CRITICAL)


def get_configuration():
    try:
        config = configparser.ConfigParser()
        config.sections()
        config.read(CONFIG_FILE_PATH)
        set_logger(config)
        return config
    except Exception:
        raise InvalidConfigException()


def get_credentials():
    try:
        with open(CREDENTIALS_FILE_PATH) as json_file:
            data = json.load(json_file)
        if data.get("username") is None and data.get("password") is None:
            raise InvalidCredentialsException(["username", "password"])
        if data.get("username") is None:
            raise InvalidCredentialsException(["username"])
        if data.get("password") is None:
            raise InvalidCredentialsException(["password"])
        return data
    except FileNotFoundError:
        raise FileNotFoundError("Error: credentials.json file not found")
    except json.JSONDecodeError:
        raise json.JSONDecodeError("Error: malformed json passed as credentials.json")


def prepare_environment(config):
    # Create download folder
    if DRY_RUN is True:
        return
    logging.debug("Prepare project environment")
    if os.path.exists(config["download"]["folder"]) is False:
        os.mkdir(config["download"]["folder"])

    if "attachments" not in os.listdir(config["download"]["folder"]):
        os.mkdir(config["download"]["folder"] + "/" + "attachments")
        logging.debug("created attachments folder")

    else:
        logging.debug("attachments folder already exists")


def gmail_connection(credentials):
    # Connect to Gmail
    logging.info("Connecting to gmail")
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(credentials["username"], credentials["password"])
    logging.info("connected correctly")
    return imap


def download_attachments(imap, config, dry_run):
    logging.info("Search emails by %s", config["imap"]["object"])
    result, email_ids = imap.search(None, '(SUBJECT "{}")'.format(config["imap"]["object"]))
    if result == 'OK':
        # Loop through the emails
        logging.info("Founded %s emails", len(email_ids[0].split()))
        for email_id in email_ids[0].split():
            result, email_body = imap.fetch(email_id, "(RFC822)")
            if result == 'OK':
                raw_email_string = email_body[0][1].decode("utf-8")
                mail = email.message_from_string(raw_email_string)
                logging.info("Parsing email with object: %s", mail['Subject'])
                for part in mail.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fileName = part.get_filename()
                    logging.info("Founded attachment %s", fileName)
                    if bool(fileName) and not dry_run:
                        filePath = os.path.join(config["download"]["folder"], 'attachments', fileName)
                        if not os.path.isfile(filePath):
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                            logging.info("Downloaded attachment %s", fileName)


def get_args(config):
    parser = argparse.ArgumentParser(description="mail-attachment-downloader help you script to search emails by object and download attachments")
    parser.add_argument("-f", "--folder_download", nargs="?", default=".", help="path folder where attachments are be downloaded")
    parser.add_argument("-d", "--dry_run", action="store_true", help="execute script without download attachments")
    args = parser.parse_args()
    config["download"]["folder"] = args.folder_download
    logging.debug(("Set download folder: %s") % (args.folder_download))
    if args.dry_run is True:
        global DRY_RUN
        DRY_RUN = True
        logging.debug("mode dry run")


def main():
    credentials = get_credentials()
    config = get_configuration()

    if credentials is not None:
        logging.debug(credentials)

    if config is not None:
        logging.debug(config)

    # Parse arguments
    get_args(config)

    prepare_environment(config)

    imap = gmail_connection(credentials)

    # Select the INBOX folder
    imap.select("INBOX")
    logging.info("Selected INBOX folder")

    download_attachments(imap, config, DRY_RUN)


if __name__ == "__main__":
    main()