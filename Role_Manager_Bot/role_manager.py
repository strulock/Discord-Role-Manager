# Regular Module Imports
from os import path, sys
# Discord API Imports
import discord
from discord.ext import commands
# Google Docs/Sheets API Imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
# Assistance Files Imports
from media import *
from request_data import *
from config import TOKEN

""" Google API Initializations """
SCOPES = link("SCOPE")
CREDENTIALS = ServiceAccountCredentials.from_json_keyfile_name(path.join(sys.path[0], "credentials.json"), SCOPES)
CLIENT = gspread.authorize(CREDENTIALS)
SERVICE = build("sheets", "v4", credentials=CREDENTIALS)

""" Discord API Initializations """
INTENTS = discord.Intents.default()
INTENTS.message_content = True
BOT = commands.Bot(command_prefix='!', intents=INTENTS)


@BOT.event
async def on_ready():
    print("BOT is ready and running!")


"""
!setuphelp - Admin Only
This command sends an embed with direct instructions on how
to get the Role Manager Bot set up and running on a server.
"""
@BOT.command()
@commands.has_permissions(administrator=True)
async def setuphelp(ctx):
    embed = discord.Embed(title="Role Manager Setup Tutorial", description="Click the link above for detailed instructions with pictures!", url=link("TUTORIAL"),  color=color("GREEN"))
    embed.add_field(name="Step 1:", value="Create a Google Sheets Worksheet.", inline=False)
    embed.add_field(name="Step 2:", value="Click the Share button on the top right and add this e-mail as an author: ```\n" + CREDENTIALS.service_account_email + "```", inline=False)
    embed.add_field(name="Step 3:", value="Select 8 columns and right-click >> Insert 8 Columns in your worksheet.", inline=False)
    embed.add_field(name="Step 4:", value="Run the !configure command and link your server with your Google spreadsheet.\n```!configure <WORKSHEET ID>```", inline=False)
    embed.add_field(name="Step 5:", value="Export the role permissions onto the Google Sheet using:\n``` !export ```", inline=False)
    embed.add_field(name="Finished!", value="The bot is now set up and you can start managing your roles! Make sure you provide a valid Spreadsheet ID, or you will encounter an error!", inline=False)
    embed.set_thumbnail(url=picture("GSHEET"))
    await ctx.send(embed=embed)


"""
!configure - Owner Only
This command creates a file for the server in the database (serverdata) 
and stores the Google Worksheet ID inside a .txt file named after the server's ID.
If a file already exists, it prompts the user to update the file instead of reconfiguring it.
"""
@BOT.command()
@commands.has_permissions(administrator=True)
async def configure(ctx, *, spreadsheet_id=None):
    if len(spreadsheet_id) == 44:  # Ensure input was given and that it is valid.
        #if ctx.message.author.id == ctx.guild.owner_id:  # If the sender is the server owner, proceed.
        file_name = str(ctx.guild.id) + ".txt"  # The name of the file is that of the server's unique ID.
        try:  # If the file exists, open and read it and give the link.
            with open(path.join("serverdata", file_name), "r+") as server_file:
                server_file.truncate(0)
                server_file.write(spreadsheet_id)

                embed = discord.Embed(title="You already have a worksheet!", description="Your spreadsheet ID has been updated instead!", color=color("GREEN"))
                embed.add_field(name="Your worksheet has been linked! Here's the link: ", value=link("SPREADSHEET") + spreadsheet_id)
                embed.set_thumbnail(url=picture("GSHEET"))
                await ctx.send(embed=embed)
        except FileNotFoundError:  # If it doesn't, create it and give the complete link.
            with open(path.join("serverdata", file_name), "w+") as server_file:
                server_file.write(spreadsheet_id)

            embed = discord.Embed(title="Worksheet Configuration Complete!", description="Your server has been added to the database.", color=color("GREEN"))
            embed.add_field(name="Your worksheet has been linked! Here's the link: ", value=link("SPREADSHEET") + spreadsheet_id)
            embed.set_thumbnail(url=picture("GSHEET"))
            await ctx.send(embed=embed)
        except Exception as exception:
            print("Server ID:" + ctx.guild.id + "\n Exception:" + str(exception))
            embed = discord.Embed(title="Something went wrong!", description="Please contact the BOT owner on GitHub!", color=color("RED"))
            embed.add_field(name="Error code: ", value=str(exception))
            embed.set_thumbnail(url=picture("ERROR"))
            await ctx.send(embed=embed)
        #else:  # If the sender is a simple Admin, refuse permission with an error embed.
        #    embed = discord.Embed(title="Access Denied!", description="You have no proper authorization for this command.", color=color("RED"))
        #    embed.add_field(name="This command may only be used by the server owner! ", value='<@' + str(ctx.guild.owner_id) + '>')
        #    embed.set_thumbnail(url=picture("ERROR"))
        #    await ctx.send(embed=embed)
    else:  # If no valid ID was given, ask for a valid ID and show instructions.
        embed = discord.Embed(title="No worksheet ID specified!", description="Please specify a valid worksheet ID.", color=color("RED"))
        embed.add_field(name="If want to see how to setup this bot use the command: ", value="```!setuphelp```", inline=False)
        embed.set_thumbnail(url=picture("ERROR"))
        await ctx.send(embed=embed)

"""
!export
This command exports all the roles and their permissions
from the Discord Server, organizes them and imports them 
to the Google Sheet assigned to that Discord Server.
"""
@BOT.command()
@commands.has_permissions(administrator=True)
async def export(ctx):
    #if ctx.message.author.id == ctx.guild.owner_id:
        file_name = str(ctx.guild.id) + ".txt"
        try:
            with open(path.join("serverdata", file_name), "r+") as server_file:
                spreadsheet_id = server_file.read()
                try:
                    role_list = ctx.guild.roles  # Export all the roles from a server. List of role type Objects.
                    role_names = [role.name for role in role_list]  # Get all the role names from the role Objects.
                    role_names.reverse() # Arrange in the order roles appear in Discord
                    role_permissions = {role: dict(role.permissions) for role in role_list}  # Put Roles in a dictionary and their permission_values in sub-dictionaries.
                    role_colors = [str(role.color) for role in role_list] # Get role colors
                    permission_names = list(role_permissions[role_list[0]].keys())  # Get all the permission names.
                    permission_names.append("Color") # Add a Color column
                    permission_values = build_rows(list(role_permissions.values()), permission_names, role_colors)  # Get all of the permissions values and convert them to √ or X.
                    permission_values.reverse() # Arrange in proper order again

                    clear_request = SERVICE.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range="A1:AH1000", body=clear_request_body())
                    titles_request = SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=titles_request_body(role_names, permission_names))
                    values_request = SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=values_request_body(permission_values))
                    clear_request.execute()  # Clears the spreadsheet.
                    titles_request.execute()
                    values_request.execute()  # Handling and execution of the requests to the Google API. See request_data.py for more info.

                    embed = discord.Embed(title="Permission Export Complete!", description="Your server's role permission_values have been successfully exported!", color=color("GREEN"))
                    embed.add_field(name="Here's the link to your worksheet: ", value=link("SPREADSHEET") + spreadsheet_id)
                    embed.set_thumbnail(url=picture("GSHEET"))
                    await ctx.send(embed=embed)
                except Exception as exception:
                    print("Server ID:" + ctx.guild.id + "\n Exception:" + str(exception))
                    embed = discord.Embed(title="Worksheet unavailable!", description="There was an issue trying to access your server's worksheet!", color=color("RED"))
                    embed.add_field(name="Make sure you have followed the !setuphelp steps correctly. If the issue persists, contact the BOT Owner.", value="```!setuphelp```")
                    embed.set_thumbnail(url=picture("ERROR"))
                    await ctx.send(embed=embed)
        except FileNotFoundError:  # If the file does not exist, prompt user to configure.
            embed = discord.Embed(title="No file found!", description="There was an issue trying to import your server's file from the database.", color=color("RED"))
            embed.add_field(name="You have to configure your server first. Please try the command !setuphelp for more information.", value="```!setuphelp```")
            embed.set_thumbnail(url=picture("ERROR"))
            await ctx.send(embed=embed)
    #else:  # If the sender is a simple Admin, refuse permission with an error embed.
        #embed = discord.Embed(title="Access Denied!", description="You have no proper authorization for this command.", color=color("RED"))
        #embed.add_field(name="This command may only be used by the server owner! ", value='<@' + str(ctx.guild.owner_id) + '>')
        #embed.set_thumbnail(url=picture("ERROR"))
        #await ctx.send(embed=embed)

"""
!imports
This command imports roles from sheet to Discord.
"""
@BOT.command()
@commands.has_permissions(administrator=True)
async def imports(ctx):
    #if ctx.message.author.id == ctx.guild.owner_id:
        file_name = str(ctx.guild.id) + ".txt"
        try:
            with open(path.join("serverdata", file_name), "r+") as server_file:
                spreadsheet_id = server_file.read()
                try:
                    values_request = SERVICE.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="A1:AQ1000")
                    valueRange = values_request.execute()
                    headings = valueRange["values"][0]  # Get headings from the first row
                    role_list = ctx.guild.roles
                    roles_to_add = []
                    for row in valueRange["values"]:
                        name = row[0]
                        if name == "":
                            continue
                        found = False
                        for role in role_list:
                            if name == role.name:
                                found = True
                                break
                        if found == False:
                            roles_to_add.append(name)
                            print("Adding role:", name)

                            #role_permissions = {role: dict(role.permissions) for role in role_list}  # Put Roles in a dictionary and their permission_values in sub-dictionaries.
                            permObject = {}
                            clr = None
                            i = 0
                            for heading in headings:
                                if heading == "":
                                    i += 1
                                    continue
                                if heading == "Color":
                                    clr = discord.Colour.from_str(row[i])
                                    i += 1
                                    continue

                                if row[i] == "✔️" or row[i] != "":
                                    permObject[heading] = True

                                i += 1

                            perms = discord.Permissions.none()
                            perms.create_instant_invite     = permObject["create_instant_invite"]                          
                            perms.kick_members              = permObject["kick_members"] 
                            perms.ban_members               = permObject["ban_members"]
                            perms.administrator             = permObject["administrator"]
                            perms.manage_channels           = permObject["manage_channels"]
                            perms.manage_guild              = permObject["manage_guild"]
                            perms.add_reactions             = permObject["add_reactions"]
                            perms.view_audit_log            = permObject["view_audit_log"]
                            perms.priority_speaker          = permObject["priority_speaker"]
                            perms.stream                    = permObject["stream"]
                            perms.read_messages             = permObject["read_messages"]
                            perms.send_messages             = permObject["send_messages"]
                            perms.send_tts_messages         = permObject["send_tts_messages"]
                            perms.manage_messages           = permObject["manage_messages"]
                            perms.embed_links               = permObject["embed_links"]
                            perms.attach_files              = permObject["attach_files"]
                            perms.read_message_history      = permObject["read_message_history"]
                            perms.mention_everyone          = permObject["mention_everyone"]
                            perms.external_emojis           = permObject["external_emojis"]
                            perms.view_guild_insights       = permObject["view_guild_insights"]
                            perms.connect                   = permObject["connect"]
                            perms.speak                     = permObject["speak"]
                            perms.mute_members              = permObject["mute_members"]
                            perms.deafen_members            = permObject["deafen_members"]
                            perms.move_members              = permObject["move_members"]
                            perms.use_voice_activation      = permObject["use_voice_activation"]
                            perms.change_nickname           = permObject["change_nickname"]
                            perms.manage_nicknames          = permObject["manage_nicknames"]
                            perms.manage_roles              = permObject["manage_roles"]
                            perms.manage_webhooks           = permObject["manage_webhooks"]
                            perms.manage_emojis             = permObject["manage_emojis"]
                            perms.use_application_commands  = permObject["use_application_commands"]
                            perms.request_to_speak          = permObject["request_to_speak"]
                            perms.manage_events             = permObject["manage_events"]
                            perms.manage_threads            = permObject["manage_threads"]
                            perms.create_public_threads     = permObject["create_public_threads"]
                            perms.create_private_threads    = permObject["create_private_threads"]
                            perms.external_stickers         = permObject["external_stickers"]
                            perms.send_messages_in_threads  = permObject["send_messages_in_threads"]
                            perms.use_embedded_activities   = permObject["use_embedded_activities"]
                            perms.moderate_members          = permObject["moderate_members"]

                            #create_role(*, name=..., permissions=..., color=..., colour=..., hoist=..., display_icon=..., mentionable=..., reason=None)
                            await ctx.guild.create_role(name=name, permissions=perms, color=clr)


                    embed = discord.Embed(title="Import Test", description="Your sheet's values, hopefully", color=color("GREEN"))
                    for role in roles_to_add:
                        embed.add_field(name="Role added:", value=role)
                    embed.set_thumbnail(url=picture("GSHEET"))
                    await ctx.send(embed=embed)

                    #role_list = ctx.guild.roles  # Export all the roles from a server. List of role type Objects.
                    #role_list.reverse()
                    #role_names = [role.name for role in role_list]  # Get all the role names from the role Objects.
                    #role_permissions = {role: dict(role.permissions) for role in role_list}  # Put Roles in a dictionary and their permission_values in sub-dictionaries.
                    #permission_names = list(role_permissions[role_list[0]].keys())  # Get all the permission names.
                    #permission_values = permission_values_to_emojis(list(role_permissions.values()), permission_names)  # Get all of the permissions values and convert them to √ or X.

                    #clear_request = SERVICE.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range="A1:AH1000", body=clear_request_body())
                    #titles_request = SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=titles_request_body(role_names, permission_names))
                    #values_request = SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=values_request_body(permission_values))
                    #clear_request.execute()  # Clears the spreadsheet.
                    #titles_request.execute()
                    #values_request.execute()  # Handling and execution of the requests to the Google API. See request_data.py for more info.

                    #embed = discord.Embed(title="Permission Export Complete!", description="Your server's role permission_values have been successfully exported!", color=color("GREEN"))
                    #embed.add_field(name="Here's the link to your worksheet: ", value=link("SPREADSHEET") + spreadsheet_id)
                    #embed.set_thumbnail(url=picture("GSHEET"))
                    #await ctx.send(embed=embed)
                except Exception as exception:
                    print("Server ID:" + ctx.guild.id + "\n Exception:" + str(exception))
                    embed = discord.Embed(title="Worksheet unavailable!", description="There was an issue trying to access your server's worksheet!", color=color("RED"))
                    embed.add_field(name="Make sure you have followed the !setuphelp steps correctly. If the issue persists, contact the BOT Owner.", value="```!setuphelp```")
                    embed.set_thumbnail(url=picture("ERROR"))
                    await ctx.send(embed=embed)
        except FileNotFoundError:  # If the file does not exist, prompt user to configure.
            embed = discord.Embed(title="No file found!", description="There was an issue trying to import your server's file from the database.", color=color("RED"))
            embed.add_field(name="You have to configure your server first. Please try the command !setuphelp for more information.", value="```!setuphelp```")
            embed.set_thumbnail(url=picture("ERROR"))
            await ctx.send(embed=embed)
    #else:  # If the sender is a simple Admin, refuse permission with an error embed.
        #embed = discord.Embed(title="Access Denied!", description="You have no proper authorization for this command.", color=color("RED"))
        #embed.add_field(name="This command may only be used by the server owner! ", value='<@' + str(ctx.guild.owner_id) + '>')
        #embed.set_thumbnail(url=picture("ERROR"))
        #await ctx.send(embed=embed)

"""
BOT RUN Command that logs in the bot with our credentials. 
Has to be in the end of the file.
"""
BOT.run(TOKEN)
