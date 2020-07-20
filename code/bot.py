from datetime import date, datetime
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
import json
import os
import requests
from requests.auth import HTTPBasicAuth
import signal
import sys
from webexteamsbot import TeamsBot
from webexteamsbot.models import Response

load_dotenv()

# Retrieve required details from environment variables
bot_email = os.getenv("PSA_BOT_EMAIL")
teams_token = os.getenv("PSA_BOT_TOKEN")
bot_url = os.getenv("PSA_BOT_URL")
bot_app_name = os.getenv("PSA_BOT_APP_NAME")
bot_scripts_api_user = os.getenv("PSA_BOT_SCRIPTS_API_USER")
bot_scripts_api_pass = os.getenv("PSA_BOT_SCRIPTS_API_PASS")


# If any of the bot environment variables are missing, terminate the app
if not bot_email or not teams_token or not bot_url or not bot_app_name:
    print(
        "sample.py - Missing Environment Variable. Please see the 'Usage'"
        " section in the README."
    )
    if not bot_email:
        print("TEAMS_BOT_EMAIL")
    if not teams_token:
        print("TEAMS_BOT_TOKEN")
    if not bot_url:
        print("TEAMS_BOT_URL")
    if not bot_app_name:
        print("TEAMS_BOT_APP_NAME")
    sys.exit()

# Create a Bot Object
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
    debug=True,
    webhook_resource_event=[
        {"resource": "messages", "event": "created"},
        {"resource": "attachmentActions", "event": "created"},
    ],
)


def create_message_with_attachment(rid, msgtxt, attachment):
    headers = {
        "content-type": "application/json; charset=utf-8",
        "authorization": "Bearer " + teams_token,
    }

    url = "https://api.ciscospark.com/v1/messages"
    data = {"roomId": rid, "attachments": [attachment], "markdown": msgtxt}
    response = requests.post(url, json=data, headers=headers)
    return response.json()


def get_attachment_actions(attachmentid):
    headers = {
        "content-type": "application/json; charset=utf-8",
        "authorization": "Bearer " + teams_token,
    }

    url = "https://api.ciscospark.com/v1/attachment/actions/" + attachmentid
    response = requests.get(url, headers=headers)
    return response.json()


def create_outlook_meeting(reminder_info):
    headers = {"Authorization": "Bearer " + o365_token}
    url = 'https://graph.microsoft.com/v1.0/me/events'
    if reminder_info["reminder_type"] == "days":
        reminder_date = date.today() + relativedelta(days=+int(reminder_info["reminder_num"]))
    elif reminder_info["reminder_type"] == "weeks":
        reminder_date = date.today() + relativedelta(weeks=+int(reminder_info["reminder_num"]))
    elif reminder_info["reminder_type"] == "months":
        reminder_date = date.today() + relativedelta(months=+int(reminder_info["reminder_num"]))
    payload = {}
    payload["subject"] = reminder_info["account"] + ": " + reminder_info["purpose"]
    startDateTime = reminder_date.strftime("%Y-%m-%d") + "T00:00:00.000"
    payload["start"] = {"dateTime": startDateTime, "timeZone": "America/New_York"}
    endDateTime = (reminder_date + relativedelta(days=+1)).strftime("%Y-%m-%d") + "T00:00:00.000"
    payload["end"] = {"dateTime": endDateTime, "timeZone": "America/New_York"}
    payload["isAllDay"] = True
    payload["showAs"] = "free"
    r = requests.post(url, json=payload, headers=headers)
    return r.status_code


def handle_cards(api, incoming_msg):
    m = get_attachment_actions(incoming_msg["data"]["id"])
    card_type = m["inputs"]["card_type"]
    if card_type == "add_reminder":
        print("Reminder info sent: ")
        print(m["inputs"])
        status_code = create_outlook_meeting(m["inputs"])
        print(status_code)
        if status_code == 201:
            return "Reminder scheduled successfully!"
        else:
            return "Error occurred during scheduling."


# Create a custom bot greeting function returned when no command is given.
# The default behavior of the bot is to return the '/help' command response
def greeting(incoming_msg):
    # Loopkup details about sender
    sender = bot.teams.people.get(incoming_msg.personId)

    # Create a Response object and craft a reply in Markdown.
    response = Response()
    response.markdown = "Hello {}, I'm a chat bot. ".format(sender.firstName)
    response.markdown += "See what I can do by asking for **/help**."
    return response


def show_reminder_card(incoming_msg):
    attachment = """
    {
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2",
            "body": [
                {
                    "type": "ColumnSet",
                    "columns": [
                        {
                            "type": "Column",
                            "width": "stretch",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "size": "Large",
                                    "text": "Add a Reminder"
                                }
                            ]
                        }
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Account",
                            "size": "Small"
                        },
                        {
                            "type": "Input.Text",
                            "placeholder": "Customer Name",
                            "id": "account"
                        }
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Purpose",
                            "size": "Small"
                        },
                        {
                            "type": "Input.Text",
                            "placeholder": "DID/SO# - Description",
                            "id": "purpose"
                        }
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Remind Me",
                            "size": "Small"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": 20,
                                    "items": [
                                        {
                                            "type": "Input.Number",
                                            "max": 999,
                                            "min": 1,
                                            "value": 30,
                                            "id": "reminder_num"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": 30,
                                    "items": [
                                        {
                                            "type": "Input.ChoiceSet",
                                            "choices": [
                                                {
                                                    "title": "Day(s)",
                                                    "value": "days"
                                                },
                                                {
                                                    "title": "Week(s)",
                                                    "value": "weeks"
                                                },
                                                {
                                                    "title": "Month(s)",
                                                    "value": "months"
                                                }
                                            ],
                                            "placeholder": "days",
                                            "value": "days",
                                            "id": "reminder_type"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": 40
                                }
                            ]
                        }
                    ]
                },
                {
                    "type": "Input.Text",
                    "isVisible": false,
                    "id": "card_type",
                    "value": "add_reminder"
                },
                {
                    "type": "ActionSet",
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Submit"
                        }
                    ]
                }
            ]
        }
    }
    """
    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(
        incoming_msg.roomId, msgtxt=backupmessage, attachment=json.loads(attachment)
    )
    print(c)
    return ""


def show_case_info_card(incoming_msg, case_info):
    ####################################################
    # REDACTED - contact bingerso@cisco.com
    ####################################################


def get_case_info(srid):
    ####################################################
    # REDACTED - contact bingerso@cisco.com
    ####################################################


def case_status(incoming_msg):
    ####################################################
    # REDACTED - contact bingerso@cisco.com
    ####################################################


# Set the bot greeting.
bot.set_greeting(greeting)

# Add commands
bot.add_command("attachmentActions", "*", handle_cards)
bot.add_command("/reminder", "Schedule a reminder", show_reminder_card)
bot.add_command("/status", "/status <SR#> to get the case status of that SR#", case_status)

# Every bot includes a default "/echo" command.  You can remove it, or any
# other command with the remove_command(command) method.
bot.remove_command("/echo")


if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=5000)