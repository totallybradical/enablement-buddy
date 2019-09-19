import os
from webexteamsbot import TeamsBot
from webexteamsbot.models import Response
import pandas as pd
import re
import sqlite3


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


def report_enablements(incoming_msg):
    """
    Sample function to do some action.
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    conn = sqlite3.connect('/home/toobradsosad/enablement-buddy/enablements.db')
    df = pd.read_sql_query("SELECT recipients, info, enablementDate FROM enablements WHERE user='" + incoming_msg.personId + "';", conn)
    # export_excel = df.to_excel (r'C:\Users\Ron\Desktop\export_dataframe.xlsx', index = None, header=True) #Don't forget to add '.xlsx' at the end of the path
    export_csv = df.to_csv ('/home/toobradsosad/enablement-buddy/exports/temp.csv', index = None, header=True) #Don't forget to add '.csv' at the end of the path
    num_enablements = df.shape[0]
    response = Response()
    response.markdown = "You've logged **" + str(num_enablements) + "** enablements! Here's a report for your records."
    response.files = "/home/toobradsosad/enablement-buddy/exports/temp.csv"
    return response


# Retrieve required details from environment variables
bot_email = os.getenv("TEAMS_BOT_EMAIL")
teams_token = os.getenv("TEAMS_BOT_TOKEN")
bot_url = os.getenv("TEAMS_BOT_URL")
bot_app_name = os.getenv("TEAMS_BOT_APP_NAME")

add_pattern = r"\/add (\d*)( *)(.+)"

# Create a Bot Object
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
    debug=True,
)

# Set the bot greeting.
bot.set_greeting(greeting)

# Add commands to the bot.
bot.add_command("/add", "Add a new enablement", add_enablement)
bot.add_command("/report", "Get a report of your enablements", report_enablements)


if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=5000)