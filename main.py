import asyncio
import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph
import time


async def main():
    print("Teams migrator\n")
    print("BE AWARE OF THROTTELING: https://learn.microsoft.com/en-us/graph/throttling-limits")
    print("SDK handles throtteling internally")

    # Load settings
    config_old = configparser.ConfigParser()
    config_old.read(["config-old-teams.cfg"])
    azure_settings_old = config_old["azure"]
    config_new = configparser.ConfigParser()
    config_new.read(["config-new-teams.cfg"])
    azure_settings_new = config_new["azure"]

    default_user = ["00000000-0000-0000-0000-000000000000", "John Doe"]

    # Key: User object ID in old tenant
    # Value: tuple of (Display name in old tenant, object ID in new tenant)
    user_map = {
        "00000000-0000-0000-0000-000000000000": (
            "John Doe",
            "00000000-0000-0000-0000-000000000000",
        ),
    }

    sharepoint_map = {
        "https://old-teams.sharepoint.com/sites/old-teams-site/Freigegebene Dokumente": "https://new-teams.sharepoint.com/sites/new-teams-site/Shared Documents",
    }

    old_teams = Graph(azure_settings_old, False, default_user, user_map, sharepoint_map)
    new_teams = Graph(azure_settings_new, True, default_user, user_map, sharepoint_map)

    teams_to_export = {
        "Old Team Name": "00000000-0000-0000-0000-000000000000",
    }
    # Mapping of old teams id to a set of channel names to export
    channels_to_export = {
        "Old Team Id": {
            "XYZ",
            "General",
        },
    }

    #new_teams_id = await new_teams.create_teams("New Teams display_name", "New Teams description")
    #print(new_teams_id)
    teams_to_import = {"New Team Name": "00000000-0000-0000-0000-000000000000"}

    try:
        await export_team(
            old_teams,
            new_teams,
            teams_to_export["Old Team Name"],
            teams_to_import["New Team Name"],
            channels_to_export,
        )
    except ODataError as odata_error:
        print("Error:")
        if odata_error.error:
            print(odata_error.error.code, odata_error.error.message)


async def export_team(
    old_teams: Graph,
    new_teams: Graph,
    old_team_id: str,
    new_team_id: str,
    channel_names: dict[str, set[str]],
):
    channels = await old_teams.list_all_channels(old_team_id)
    for channel in channels:
        if channel.display_name not in channel_names[old_team_id]:
            print(f"skipping channel: {channel.display_name} {channel.id}")
            continue
        print(f"work on channel: {channel.display_name} {channel.id}")
        new_channel = await new_teams.get_channel(new_team_id, channel.display_name)
        if new_channel is None:
            new_channel = await new_teams.create_channel(new_team_id, channel)
        print(f"new channel: {new_channel.display_name} {new_channel.id}")
        channel_members = await old_teams.list_channel_members(old_team_id, channel.id)
        print("Channel members:")
        for channel_member in channel_members:
            print(channel_member.display_name)
        messages = await old_teams.list_messages(old_team_id, channel.id)
        print(f"Obtained {len(messages)} old messages")
        for message in messages:
            if message.deleted_date_time is not None:
                continue
            new_msg = await new_teams.send_message(new_team_id, new_channel.id, message)
            if new_msg.id == "Msg already exists":
                print(new_msg.id)
                continue
            print(f"Msg {new_msg.id} sent to channel {new_channel.id} in teams {new_team_id}")
            replies = await old_teams.list_replies(old_team_id, channel.id, message.id)
            print(f"Obtained {len(replies)} old replies")
            for reply in replies:
                if reply.deleted_date_time is not None:
                    continue
                new_reply = await new_teams.send_reply(
                    new_team_id, new_channel.id, new_msg.id, reply
                )
                print(
                    f"Replied {new_reply.id} to msg {new_msg.id} on channel {new_channel.id} in teams {new_team_id}"
                )
        await new_teams.complete_channel_migration(new_team_id, new_channel.id)
    print("all channels migrated")
    time.sleep(10)
    await new_teams.complete_teams_migration(new_team_id)
    print("migration finished")
    await new_teams.add_teams_member(new_team_id, new_teams.default_user[0])
    print(f"{new_teams.default_user[1]} added as Teams owner")


# Run main
asyncio.run(main())
