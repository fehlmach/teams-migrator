import requests
import os
import re
from configparser import SectionProxy
from msal import PublicClientApplication
from azure.identity import InteractiveBrowserCredential
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
from msgraph.generated.teams.item.channels.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.models.directory_object import DirectoryObject
from msgraph.generated.models.chat_message_attachment import ChatMessageAttachment
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.chat_message_reaction import ChatMessageReaction
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.channel import Channel
from msgraph.generated.models.group import Group
from msgraph.generated.models.conversation_member import ConversationMember
from msgraph.generated.models.chat_message_from_identity_set import (
    ChatMessageFromIdentitySet,
)

from msgraph.generated.models.chat_message_mention import ChatMessageMention
from msgraph.generated.teams.item.channels.channels_request_builder import (
    ChannelsRequestBuilder,
)
from msgraph.generated.models.o_data_errors.o_data_error import ODataError


class Graph:
    settings: SectionProxy
    app: PublicClientApplication
    graph_scopes: list[str]
    credential: InteractiveBrowserCredential
    client: GraphServiceClient
    default_user: list[str]
    user_map: dict[str, str]
    sharepoint_map: dict[str, str]
    tenant_id: str

    def __init__(
        self,
        config: SectionProxy,
        is_client_credential: bool,
        default_user: list[str],
        user_map: dict[str, str],
        sharepoint_map: dict[str, str],
    ):
        self.settings = config
        client_id = self.settings["clientId"]
        self.tenant_id = self.settings["tenantId"]
        self.graph_scopes = self.settings["graphUserScopes"].split(" ")
        self.default_user = default_user
        self.user_map = user_map
        self.sharepoint_map = sharepoint_map

        if is_client_credential:
            self.credential = ClientSecretCredential(
                self.tenant_id,
                client_id,
                os.environ.get("CLIENT_SECRET"),
            )
            self.client = GraphServiceClient(credentials=self.credential, scopes=self.graph_scopes)
        else:
            self.app = PublicClientApplication(
                client_id, authority=f"https://login.microsoftonline.com/{self.tenant_id}"
            )
            self.credential = InteractiveBrowserCredential(
                client_id=client_id, tenant_id=self.tenant_id
            )
            self.client = GraphServiceClient(self.credential, self.graph_scopes)

    async def get_user_token(self):
        result = self.credential.get_token("User.Read")
        return result.token

    # https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0&tabs=python
    async def list_group_membership(self, group_id: str) -> list[DirectoryObject]:
        members = await self.client.groups.by_group_id(group_id).members.get()
        return members.value

    # https://learn.microsoft.com/en-us/graph/teams-list-all-teams
    # https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=python
    async def list_teams(self) -> list[Group]:
        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            filter="resourceProvisioningOptions/Any(x:x eq 'Team')",
        )
        request_configuration = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
        )
        # request_configuration.headers.add("ConsistencyLevel", "eventual")
        teams = await self.client.groups.get(request_configuration=request_configuration)
        return teams.value

    # https://learn.microsoft.com/en-us/graph/api/team-post?view=graph-rest-beta&tabs=python&preserve-view=true
    async def create_teams(
        self,
        display_name: str,
        description: str,
    ) -> str:
        # Using the SDK does not work, as the additional_data field does not get parsed into the request.
        # request_body = Team(
        #     display_name=display_name,
        #     description=display_name,
        #     created_date_time="2015-01-01T11:11:11.111Z",
        #     additional_data={
        #         "@microsoft_graph_team_creation_mode": "migration",
        #         "template@odata_bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
        #     },
        # )
        # teams = await self.client.teams.post(request_body)
        url = "https://graph.microsoft.com/v1.0/teams/"
        json_body = {
            "@microsoft.graph.teamCreationMode": "migration",
            "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
            "displayName": display_name,
            "description": description,
            "createdDateTime": "2015-01-01T11:11:11.000Z",
        }
        access = await self.credential.get_token(self.graph_scopes[0])
        headers = {
            "Authorization": "Bearer " + access.token,
            "Content-Type": "application/json",
        }
        response = requests.post(url, headers=headers, json=json_body)
        if response.status_code == 202:
            print("Teams created successfully.")
        else:
            print("Error creating teams. Status code:", response.status_code)
            print("Response:", response.text)
        pattern = r"'([a-f0-9-]+)'"
        match = re.search(pattern, response.headers.get("Content-Location"))
        return match.group(1)

    # https://learn.microsoft.com/en-us/graph/api/channel-list?view=graph-rest-1.0&tabs=python
    async def list_all_channels(self, team_id: str) -> list[Channel]:
        channels = await self.client.teams.by_team_id(team_id).channels.get()
        return channels.value

    # https://learn.microsoft.com/en-us/graph/api/channel-list?view=graph-rest-1.0&tabs=python
    async def get_channel(self, team_id: str, channel_name: str) -> Channel | None:
        query_params = ChannelsRequestBuilder.ChannelsRequestBuilderGetQueryParameters(
            filter=f"displayName eq '{channel_name}'",
        )
        request_configuration = (
            ChannelsRequestBuilder.ChannelsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
            )
        )
        result = await self.client.teams.by_team_id(team_id).channels.get(
            request_configuration=request_configuration
        )
        return result.value[0] if len(result.value) > 0 else None

    async def create_channel(self, team_id: str, old_channel: Channel) -> Channel:
        # Using the SDK does not work, as the additional_data field does not get parsed into the request.
        # request_body = Channel(
        #     display_name=old_channel.display_name,
        #     description=old_channel.description,
        #     membership_type=ChannelMembershipType.Standard,
        #     additional_data={
        #         "@microsoft_graph_channel_creation_mode": "migration",
        #     },
        # )
        # new_channel = await self.client.teams.by_team_id(team_id).channels.post(request_body)
        # return new_channel
        url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels"
        json_body = {
            "@microsoft.graph.channelCreationMode": "migration",
            "displayName": old_channel.display_name,
            "description": old_channel.description,
            "membershipType": "standard",
            "createdDateTime": old_channel.created_date_time.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]
            + "Z",
        }
        bearer_token = await self.credential.get_token(self.graph_scopes[0])
        headers = {
            "Authorization": "Bearer " + bearer_token.token,
            "Content-Type": "application/json",
        }
        response = requests.post(url, headers=headers, json=json_body)
        if response.status_code == 201:
            print("Channel created successfully.")
        else:
            print("Error creating channel. Status code:", response.status_code)
            print("Response:", response.text)

        return await self.get_channel(team_id, old_channel.display_name)

    # https://learn.microsoft.com/en-us/graph/api/channel-list-members?view=graph-rest-1.0&tabs=python
    async def list_channel_members(self, team_id: str, channel_id: str) -> list[ConversationMember]:
        members = (
            await self.client.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .members.get()
        )
        return members.value

    # https://learn.microsoft.com/en-us/graph/api/team-post-members?view=graph-rest-1.0&tabs=python
    async def add_teams_member(self, team_id: str, user_id: str):
        # Using the SDK does not work, as the additional_data field does not get parsed into the request.
        # request_body = AadUserConversationMember(
        #     odata_type="#microsoft.graph.aadUserConversationMember",
        #     roles=["owner"],
        #     additional_data={
        #         "user@odata_bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')",
        #     },
        # )
        # await self.client.teams.by_team_id(team_id).members.post(request_body)
        url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/members"
        json_body = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/beta/users/{user_id}",
        }
        bearer_token = await self.credential.get_token(self.graph_scopes[0])
        headers = {
            "Authorization": "Bearer " + bearer_token.token,
            "Content-Type": "application/json",
        }
        response = requests.post(url, headers=headers, json=json_body)
        if response.status_code == 201:
            print("Member added successfully.")
        else:
            print("Error adding member to teams. Status code:", response.status_code)
            print("Response:", response.text)

    # https://learn.microsoft.com/en-us/graph/api/channel-list-messages?view=graph-rest-1.0&tabs=python
    async def list_messages(self, team_id: str, channel_id: str) -> list[ChatMessage]:
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=50,
        )
        request_configuration = (
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
            )
        )
        messages = (
            await self.client.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .messages.get(request_configuration=request_configuration)
        )
        chat_messages = messages.value
        next_link = messages.odata_next_link
        while next_link:
            messages = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.with_url(next_link)
                .get()
            )
            next_link = messages.odata_next_link
            chat_messages.extend(messages.value)
        return [msg for msg in chat_messages if msg.message_type == "message"]

    # https://learn.microsoft.com/en-us/graph/api/channel-post-messages?view=graph-rest-1.0&tabs=python
    async def send_message(
        self, team_id: str, channel_id: str, old_msg: ChatMessage
    ) -> ChatMessage:
        request_body = ChatMessage(
            message_type=old_msg.message_type,
            created_date_time=old_msg.created_date_time,
            subject=old_msg.subject,
            summary=old_msg.summary,
            from_=old_msg.from_,
            body=old_msg.body,
            attachments=self.map_attachments(old_msg.attachments),
            mentions=old_msg.mentions,
        )
        self.map_user(request_body.from_)
        self.map_mentions_user(request_body.mentions)
        self.replace_inexistant_users(request_body)
        # It is not yet possible to import reactions...
        # for i in range(len(request_body.reactions)):
        #    request_body.reactions[i].user = self.map_user(request_body.reactions[i].user)
        self.add_reaction_to_body(request_body.body, old_msg.reactions)
        request_body.body.content = request_body.body.content.replace("&nbsp;", " ")
        # Reformat emojis
        if "<emoji id=" in request_body.body.content:
            request_body.body.content = request_body.body.content.replace('" title=""></emoji>', "")
            pattern = r'<emoji id="[a-zA-Z]+" alt="'
            request_body.body.content = re.sub(pattern, "", request_body.body.content)
        print("reply requestbody constructed: %s", request_body)
        try:
            return (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.post(request_body)
            )
        except ODataError as odata_error:
            if odata_error.response_status_code == 409:
                return ChatMessage(id="Msg already exists")
            raise odata_error

    # https://learn.microsoft.com/en-us/graph/api/chatmessage-list-replies?view=graph-rest-1.0&tabs=python
    async def list_replies(
        self, team_id: str, channel_id: str, chat_message_id: str
    ) -> list[ChatMessage]:
        replies = (
            await self.client.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .messages.by_chat_message_id(chat_message_id)
            .replies.get()
        )
        reply_messages = replies.value
        next_link = replies.odata_next_link
        while next_link:
            replies = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(chat_message_id)
                .replies.with_url(next_link)
                .get()
            )
            next_link = replies.odata_next_link
            reply_messages.extend(replies.value)
        return reply_messages

    # https://learn.microsoft.com/en-us/graph/api/chatmessage-post-replies?view=graph-rest-1.0&tabs=python
    async def send_reply(
        self, team_id: str, channel_id: str, chat_message_id: str, old_reply: ChatMessage
    ) -> ChatMessage:
        request_body = ChatMessage(
            message_type=old_reply.message_type,
            created_date_time=old_reply.created_date_time,
            subject=old_reply.subject,
            summary=old_reply.summary,
            from_=old_reply.from_,
            body=old_reply.body,
            attachments=self.map_attachments(old_reply.attachments),
            mentions=old_reply.mentions,
        )
        self.map_user(request_body.from_)
        self.map_mentions_user(request_body.mentions)
        self.replace_inexistant_users(request_body)
        # It is not yet possible to import reactions...
        # for i in range(len(request_body.reactions)):
        #    request_body.reactions[i].user = self.map_user(request_body.reactions[i].user)
        self.add_reaction_to_body(request_body.body, old_reply.reactions)
        request_body.body.content = request_body.body.content.replace("&nbsp;", " ")
        # Reformat emojis
        if "<emoji id=" in request_body.body.content:
            print("working on emoji regex")
            request_body.body.content = request_body.body.content.replace('" title=""></emoji>', "")
            pattern = r'<emoji id="[a-zA-Z]+" alt="'
            request_body.body.content = re.sub(pattern, "", request_body.body.content)
        print("reply requestbody constructed: %s", request_body)
        return (
            await self.client.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .messages.by_chat_message_id(chat_message_id)
            .replies.post(request_body)
        )

    # https://learn.microsoft.com/en-us/graph/api/channel-completemigration?view=graph-rest-1.0&tabs=python
    async def complete_channel_migration(self, team_id: str, channel_id: str):
        await self.client.teams.by_team_id(team_id).channels.by_channel_id(
            channel_id
        ).complete_migration.post()

    # https://learn.microsoft.com/en-us/graph/api/team-completemigration?view=graph-rest-1.0&tabs=python
    async def complete_teams_migration(self, team_id: str):
        await self.client.teams.by_team_id(team_id).complete_migration.post()

    def map_attachments(
        self, attachments: list[ChatMessageAttachment]
    ) -> list[ChatMessageAttachment]:
        print("working on attachments")
        new_attachments = []
        for attachment in attachments:
            if attachment.content_type != "reference":
                new_attachments.append(attachment)
                continue
            was_replaced = False
            for key, value in self.sharepoint_map.items():
                if attachment.content_url.startswith(key):
                    end = attachment.content_url[len(key) :]
                    id = attachment.id
                    name = attachment.name
                    new_attachments.append(
                        ChatMessageAttachment(
                            content_type="reference",
                            content_url=value + end,
                            id=id,
                            name=name,
                        )
                    )
                    was_replaced = True
                    break
            if not was_replaced:
                print("Didn't replace attachment: %s", attachment.content_url)
        return new_attachments

    def map_user(self, sender: ChatMessageFromIdentitySet):
        if sender.user.id in self.user_map.keys():
            sender.user.id = self.user_map[sender.user.id][1]
        sender.user.additional_data["tenantId"] = self.tenant_id

    def map_mentions_user(self, mentions: list[ChatMessageMention]):
        for i in range(len(mentions)):
            if mentions[i].mentioned.user is not None:
                self.map_user(mentions[i].mentioned)

    def add_reaction_to_body(self, body: ItemBody, reactions: list[ChatMessageReaction]):
        if not reactions:
            return
        if body.content_type.value != "html":
            body.content_type = BodyType["Html"]
            body.content = f"<div>{body.content}</div>"
        body.content = f"{body.content}\n-----\n"
        for reaction in reactions:
            display_name = ""
            if reaction.user.user.id in self.user_map:
                display_name = self.user_map[reaction.user.user.id][0]
            elif reaction.user.user.display_name is not None:
                display_name = reaction.user.user.display_name
            else:
                display_name = reaction.user.user.id
            body.content = (
                f"{body.content}<p>{display_name}'s Reaktion: {reaction.reaction_type}</p>\n"
            )
        body.content = f"{body.content}-----"

    def replace_inexistant_users(self, request: ChatMessage):
        if request.from_.user.id in self.user_map.keys():
            return
        if request.body.content_type.value != "html":
            request.body.content_type = BodyType["Html"]
            request.body.content = f"<div>{request.body.content}</div>"
        request.body.content = f"\n<p>-----</p>\n<b>Original message from: {request.from_.user.display_name}</b>\n<p>-----</p>\n{request.body.content}"
        request.from_.user.id = self.default_user[0]
        request.from_.user.display_name = self.default_user[1]
