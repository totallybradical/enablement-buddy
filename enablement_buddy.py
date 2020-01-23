from dotenv import load_dotenv
from datetime import datetime
from datetime import date
import json
import os
from webexteamsbot import TeamsBot
from webexteamsbot.models import Response
import pandas as pd
import re
import requests
import sqlite3


load_dotenv()

# Retrieve required details from environment variables
bot_email = os.getenv("TEAMS_BOT_EMAIL")
teams_token = os.getenv("TEAMS_BOT_TOKEN")
bot_url = os.getenv("TEAMS_BOT_URL")
bot_app_name = os.getenv("TEAMS_BOT_APP_NAME")

add_pattern = r".*\/add (\d*)( *)(.+)"

# Create a Bot Object
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
    webhook_resource_event=[{"resource": "messages", "event": "created"},
                            {"resource": "attachmentActions", "event": "created"}]
)


# Create a custom bot greeting function returned when no command is given.
# The default behavior of the bot is to return the '/help' command response
def greeting(incoming_msg):
    # Lookup details about sender
    sender = bot.teams.people.get(incoming_msg.personId)

    # Create a Response object and craft a reply in Markdown.
    response = Response()
    response.markdown = "Hello {}, I'm Enablement Buddy! ".format(sender.firstName)
    response.markdown += "See what I can do by asking for **/help**."
    return response


def add_enablement(incoming_msg):
    """
    Sample function to do some action.
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    extraction = re.match(add_pattern, incoming_msg.text)
    if not extraction:
        return "Invalid entry - please write /add <OPTIONAL number of recipients> <description>"
    qty = extraction.group(1)
    desc = extraction.group(3)

    if not desc:
        return "Invalid entry - did you include a description?"

    conn = sqlite3.connect('/home/toobradsosad/enablement-buddy/enablements.db')
    c = conn.cursor()

    # enablements(id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT NOT NULL, recipients INTEGER DEFAULT(1), info TEXT, enablementDate DATETIME DEFAULT(getdate()));
    if qty:
        c.execute("INSERT INTO enablements (user, recipients, info) VALUES ('" + incoming_msg.personId + "', " + qty + ", '" + desc + "');")
    else:
        c.execute("INSERT INTO enablements (user, info) VALUES ('" + incoming_msg.personId + "', '" + desc + "');")
    
    conn.commit()
    conn.close()

    return "Enablement added successfully!"


# This function generates a basic adaptive card and sends it to the user
# You can use Microsofts Adaptive Card designer here:
# https://adaptivecards.io/designer/. The formatting that Webex Teams
# uses isn't the same, but this still helps with the overall layout
# make sure to take the data that comes out of the MS card designer and
# put it inside of the "content" below, otherwise Webex won't understand
# what you send it.
def show_card(incoming_msg):
    today = date.today()
    today_str = today.strftime('%m/%d/%Y')
    attachment = '''
    {
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "ColumnSet",
                    "columns": [
                        {
                            "type": "Column",
                            "width": 2,
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "Add an Activity",
                                    "weight": "Bolder",
                                    "size": "Medium"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Log an activity that you've spent time on.",
                                    "isSubtle": true,
                                    "wrap": true
                                },
                                {
                                    "type": "TextBlock",
                                    "size": "Small",
                                    "text": "Type"
                                },
                                {
                                    "type": "Input.ChoiceSet",
                                    "id": "activity_type",
                                    "placeholder": "Choose an activity type...",
                                    "choices": [
                                        {
                                            "title": "Enablement",
                                            "value": "enablement"
                                        },
                                        {
                                            "title": "Post-Sales",
                                            "value": "postsales"
                                        }
                                    ]
                                },
                                {
                                    "type": "TextBlock",
                                    "size": "Small",
                                    "text": "Date"
                                },
                                {
                                    "type": "Input.Date",
                                    "value": "''' + today_str + '''",
                                    "id": "date"
                                },
                                {
                                    "type": "TextBlock",
                                    "size": "Small",
                                    "text": "Duration (hours)"
                                },
                                {
                                    "type": "TextBlock",
                                    "size": "Small",
                                    "text": "Description"
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "description",
                                    "placeholder": "What did you do?"
                                }
                            ]
                        }
                    ]
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit"
                }
            ]
        }
    }
    '''
    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(incoming_msg.roomId,
                                       msgtxt=backupmessage,
                                       attachment=json.loads(attachment))
    return ""


# An example of how to process card actions
def handle_cards(api, incoming_msg):
    """
    Sample function to handle card actions.
    :param api: webexteamssdk object
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    m = get_attachment_actions(incoming_msg["data"]["id"])
    activity_type = m["inputs"]["activity_type"]
    description = m["inputs"]["description"]
    duration = m["inputs"]["duration"]
    date = m["inputs"]["date"]
    date_object = datetime.strptime(date, '%m/%d/%Y')
    date_str = date_object.strftime('%Y-%m-%d')
    conn = sqlite3.connect('/home/toobradsosad/enablement-buddy/tracking.db')
    c = conn.cursor()

    # enablements(id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT NOT NULL, recipients INTEGER DEFAULT(1), info TEXT, enablementDate DATETIME DEFAULT(getdate()));
    c.execute("INSERT INTO tracking (user, description, duration, activityDate) VALUES ('" + incoming_msg["actorId"] + "', '" + description + "', '" + duration + "', '" + date_str + "');")

    conn.commit()
    conn.close()

    return "Added successfully!"


# Temporary function to send a message with a card attachment (not yet
# supported by webexteamssdk, but there are open PRs to add this
# functionality)
def create_message_with_attachment(rid, msgtxt, attachment):
    headers = {
        'content-type': 'application/json; charset=utf-8',
        'authorization': 'Bearer ' + teams_token
    }

    url = 'https://api.ciscospark.com/v1/messages'
    data = {"roomId": rid, "attachments": [attachment], "markdown": msgtxt}
    response = requests.post(url, json=data, headers=headers)
    return response.json()


# Temporary function to get card attachment actions (not yet supported
# by webexteamssdk, but there are open PRs to add this functionality)
def get_attachment_actions(attachmentid):
    headers = {
        'content-type': 'application/json; charset=utf-8',
        'authorization': 'Bearer ' + teams_token
    }

    url = 'https://api.ciscospark.com/v1/attachment/actions/' + attachmentid
    response = requests.get(url, headers=headers)
    return response.json()


def generate_report(incoming_msg):
    """
    Sample function to do some action.
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    conn = sqlite3.connect('/home/toobradsosad/enablement-buddy/tracking.db')
    df = pd.read_sql_query("SELECT description, activityDate FROM tracking WHERE user='" + incoming_msg.personId + "';", conn)
    export_excel = df.to_excel ('/home/toobradsosad/enablement-buddy/exports/activities.xlsx', index = None, header=True) #Don't forget to add '.xlsx' at the end of the path
    # export_csv = df.to_csv('/home/toobradsosad/enablement-buddy/exports/temp.csv', index = None, header=True) #Don't forget to add '.csv' at the end of the path
    num_enablements = df.shape[0]
    response = Response()
    response.markdown = "You've logged **" + str(num_enablements) + "** activities! Here's a report for your records."
    response.files = "/home/toobradsosad/enablement-buddy/exports/activities.xlsx"
    return response


# Set the bot greeting.
bot.set_greeting(greeting)

# Add commands to the bot.
bot.add_command('attachmentActions', '*', handle_cards)
bot.add_command("/add", "Add a new enablement (/add <description> OR /add <# recipients> <description>)", add_enablement)
bot.add_command("/card", "Log an activity using a card!", show_card)
bot.add_command("/report", "Get a report of your activities to-date.", generate_report)
bot.remove_command("/echo")

if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=5000)