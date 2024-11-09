import argparse
import asyncio
import base64
from functools import cache
import glob
import json
import os
import pprint
import re
import urllib
import sys
from typing import Any, Dict, Generator, Optional, overload

from azure.identity import InteractiveBrowserCredential
import dateparser
from jinja2 import Environment, FileSystemLoader
from kiota_abstractions.native_response_handler import NativeResponseHandler
from kiota_http.middleware.options import ResponseHandlerOption
import kiota_serialization_json 
import kiota_serialization_json.json_parse_node_factory
import kiota_serialization_json.json_serialization_writer_factory
from msgraph import GraphServiceClient
from msgraph.generated.models.chat import Chat
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.chat_message_attachment import ChatMessageAttachment
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from msgraph.generated.users.item.chats.item.messages.messages_request_builder import MessagesRequestBuilder
from kiota_abstractions.api_error import APIError
import pytz
import datetime

client_id = os.getenv("CLIENT_ID")

filename_size_limit = 255


def makedir(path: str):
    """basically mkdir -p"""
    if not os.path.exists(path):
        os.makedirs(os.path.join(path), exist_ok=True)


@cache
def get_jinja_env():
    jinja_env = Environment(loader=FileSystemLoader("templates"))
    jinja_env.filters["localdt"] = localdt # type: ignore
    return jinja_env


def localdt(dt: datetime.datetime | None, format: str ="%m/%d/%Y %I:%M %p %Z"):
    """parse a date string into a datetime object, localize it, and format it for display"""
    if dt is None:
        return ""
    tz = pytz.timezone("Europe/London")
    local_dt = dt.astimezone(tz)
    return local_dt.strftime(format)


def get_member_list(chat: Chat):
    """return a sorted comma-separated list of chat members"""
    if chat.members is None:
        return "No Members"
    members = [
        m.display_name if m.display_name else "No Name" for m in chat.members
    ]
    return ", ".join(sorted(members))

def sanitize_filename(filename: str, allow_unicode: bool = True) -> str:
    """sanitize a filename to be suitable for Windows and Unix filesystems.
    this function is idempotent.
    it will remove e.g. slashes and colons but spaces, dots etc are allowed."""
    pattern = r"[^-\w.,()]"
    if allow_unicode:
        pattern = r"(?u)" + pattern
    return re.sub(pattern, "", filename.strip())


def get_chat_name(chat: Chat):
    """get a "name" for the chat: either its topic or a comma-separated list of members"""
    if chat.topic:
        name = chat.topic
    else:
        name = get_member_list(chat)
    return name


def get_hosted_content_filename(msg_id: str, hosted_content_id: str):
    """
    return a base filename for the msg_id + hosted_content_id,
    truncating if necessary to keep it under the filename size limit
    """
    filename = f"hosted_content_{msg_id}_{hosted_content_id}"
    return sanitize_filename(filename[0:filename_size_limit])


def get_hosted_content_id(attachment: ChatMessageAttachment) -> str:
    """extract the hosted_content_id from the Attachment dict record"""
    # it's stupid that I have to parse this. codeSnippetUrl already is the complete URL
    # but I can't figure out how to make a request to it directly using the client object
    if not attachment.content:
        return ""
    content = json.loads(attachment.content)
    hosted_content_id = content["codeSnippetUrl"].split("/")[-2]
    return hosted_content_id

@overload
async def fetch_all_for_request(getable: ChatsRequestBuilder, request_config: ChatsRequestBuilder.ChatsRequestBuilderGetRequestConfiguration) -> Generator[Chat, Any, None]:
    ...

@overload
async def fetch_all_for_request(getable: MessagesRequestBuilder, request_config: MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration) -> Generator[ChatMessage, Any, None]:
    ...

async def fetch_all_for_request(getable: ChatsRequestBuilder | MessagesRequestBuilder, request_config: ChatsRequestBuilder.ChatsRequestBuilderGetRequestConfiguration | MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration):
    """
    returns an iterator over the dict records returned from a request

    getable = an MS Graph API object with a get() method.
    request_config = request configuration object to pass to get()
    """
    target_request = getable
    while target_request:
        response = await target_request.get(request_configuration=request_config) # type: ignore
        if not response:
            print("  Error: no response")
            break

        target_request = getable.with_url(response.odata_next_link) if response.odata_next_link else None

        if not response.value:
            print("  Error: no response.value")
            break
        for result in response.value:
            yield result


async def download_hosted_content(
    client: GraphServiceClient, chat: Chat, msg: ChatMessage, hosted_content_id: str, chat_dir: str
):
    if not chat.id or not msg.id:
        print("  Error: chat or msg id is None")
        return

    # it's happened in one case that a user doesn't have access to the hosted content
    # in a chat they're a member of. not sure how that's possible, but that's why
    # this check is here.
    try:
        result = (
            await client.chats.by_chat_id(chat.id)
            .messages.by_chat_message_id(msg.id)
            .hosted_contents.by_chat_message_hosted_content_id(hosted_content_id)
            .content.get()
        )
    except Exception as e:
        print("  Error: " + str(e))
        result = str(e).encode()
    filename = get_hosted_content_filename(msg.id, hosted_content_id)
    path = os.path.join(chat_dir, filename)
    with open(path, "wb") as f:
        f.write(result)


async def download_sharepoint_document(client: GraphServiceClient, url: str, chat_dir: str):
    """
    Download a file from a SharePoint URL using the Graph API.
    """
    # Extract the site name and path from the URL
    match = re.match(
        r"https://[a-z0-9-]+\.sharepoint\.com/personal/([^/]+)/Documents/(.+)", url
    )
    if not match:
        print(f"Error: URL does not match expected format: {url}")
        return

    user, file_path = match.groups()
    file_path = urllib.parse.unquote(file_path)

    try:
        # Download the file content
        b64url = base64.b64encode(url.encode()).decode()
        b64url = "u!" + b64url.replace("/", "_").replace("+", "-").replace("=", "")
        content_bytes = await client.shares.by_shared_drive_item_id(b64url).drive_item.content.get()
        # Save the file
        path = os.path.join(chat_dir, user, file_path)
        makedir(os.path.dirname(path))
        with open(path, "wb") as f:
            f.write(content_bytes)
    except APIError as e:
        # If the sharing link no longer exists, just warn and continue, can't do
        # anything about it
        if "error" in e and "The sharing link no longer exists" in e.error.message:
            print(f"Warn: The sharing link for {url} no longer exists")
        else:
            raise e
    except Exception as e:
        print(f"Error downloading file from URL {url}: {str(e)}")
        exit(1)


async def download_hosted_content_in_msg(client: GraphServiceClient, chat: Chat, msg: ChatMessage, chat_dir: str):
    # fetch all the "hosted contents" (inline attachments)
    if not msg.attachments:
        return
     
    for attachment in msg.attachments:
        if attachment.content_type == "application/vnd.microsoft.card.codesnippet":
            hosted_content_id = get_hosted_content_id(attachment)
            await download_hosted_content(
                client, chat, msg, hosted_content_id, chat_dir
            )
        elif attachment.content_type == "reference" and attachment.content_url:
            # Download referenced attachments by URL too, as i hate SharePoint
            url = attachment.content_url
            matches = re.findall(r"https://([a-z0-9-]+)\.sharepoint\.com", url)
            if matches:
                await download_sharepoint_document(client, url, chat_dir)

    # images are not present as attachments, just referenced in img tags
    content_type = msg.body.content_type if msg.body and msg.body.content_type else ""
    content = msg.body.content if msg.body and msg.body.content else ""
    if content_type == "html":
        for match in re.findall('src="(.+?)"', content):
            url = match
            if "https://graph.microsoft.com/v1.0/chats/" in url:
                hosted_content_id = url.split("/")[-2]
                await download_hosted_content(
                    client, chat, msg, hosted_content_id, chat_dir
                )


async def download_messages(client: GraphServiceClient, chat: Chat, chat_dir: str, force: bool = False):
    """
    download messages for a chat, including its 'hosted content'

    Note that msg ids are not globally unique. They're millisecond timestamps.

    the 'force' flag downloads all messages that haven't been saved yet.
    by default, only newer messages are downloaded.
    """
    if chat.id is None:
        return

    async def save_msg(msg: ChatMessage):
        kiota_factory = kiota_serialization_json.json_serialization_writer_factory.JsonSerializationWriterFactory()
        kiota_writer = kiota_factory.get_serialization_writer(kiota_factory.get_valid_content_type())
        msg.serialize(kiota_writer)

        with open(path, "wb") as f:
            f.write(kiota_writer.get_serialized_content())
        await download_hosted_content_in_msg(client, chat, msg, chat_dir)

    last_msg_id = chat.last_message_preview.id if chat.last_message_preview is not None else None
    last_msg_exists = os.path.exists(os.path.join(chat_dir, sanitize_filename(f"msg_{last_msg_id}.json")))
    if force or not last_msg_id or not last_msg_exists:
        count_saved = 0
        count_updated = 0
        count_unchanged = 0
        messages_request = client.me.chats.by_chat_id(chat.id).messages

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=50,
        )
        request_config = (
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
            )
        )

        async for msg in fetch_all_for_request(messages_request, request_config):
            path = os.path.join(chat_dir, sanitize_filename(f"msg_{msg.id}.json"))
            if not os.path.exists(path):
                await save_msg(msg)
                count_saved += 1
            else:
                # if incoming msg was deleted, we don't want to overwrite our file
                if not msg.deleted_date_time:
                    with open(path, "r") as f:
                        existing_msg = json.loads(f.read())

                    # save edited/modified msgs
                    if (
                        existing_msg["lastModifiedDateTime"]
                        != msg.last_modified_date_time
                        or existing_msg["lastEditedDateTime"]
                        != msg.last_edited_date_time
                    ):
                        await save_msg(msg)
                        count_updated += 1
                    else:
                        count_unchanged += 1
                        # msg exists but hasn't been edited/modified, so we can stop
                        # if we're not running in force mode
                        if not force:
                            break
                else:
                    count_unchanged += 1

        output = f"  Message counts: {count_saved} saved, {count_updated} updated"
        if force:
            output += f", {count_unchanged} unchanged"
        print(output)
    else:
        print("  No new messages in the chat since last run")


async def download_chat(client: GraphServiceClient, chat: Chat, data_dir: str, force: bool):
    """download a single chat and its associated data (messages, attachments)"""
    print(f"Processing chat {get_chat_name(chat)} (id {chat.id})")

    if chat.id is None:
        print("  Skipping chat with no id")
        return

    chat_dir = os.path.join(data_dir, sanitize_filename(chat.id))
    makedir(chat_dir)

    kiota_factory = kiota_serialization_json.json_serialization_writer_factory.JsonSerializationWriterFactory()
    kiota_writer = kiota_factory.get_serialization_writer(kiota_factory.get_valid_content_type())
    chat.serialize(kiota_writer)

    with open(f"{chat_dir}.json", "wb") as f:
        f.write(kiota_writer.get_serialized_content())

    await download_messages(client, chat, chat_dir, force)


async def download_all(output_dir: str, force: bool):
    """download all chats"""
    data_dir = os.path.join(output_dir, "data")
    makedir(data_dir)

    client = get_graph_client()

    print("Opening browser window for authentication")

    query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        expand=["members", "lastMessagePreview"], top=50
    )
    request_config = ChatsRequestBuilder.ChatsRequestBuilderGetRequestConfiguration(
        query_parameters=query_params,
    )
    async for chat in fetch_all_for_request(client.me.chats, request_config):
        await download_chat(client, chat, data_dir, force)


def render_hosted_content(msg: ChatMessage, hosted_content_id: str, chat_dir: str):
    if not msg.id:
        return "Error: no msg id" 
    filename = get_hosted_content_filename(msg.id, hosted_content_id)
    path = os.path.join(chat_dir, filename)
    with open(path, "r") as f:
        data = f.read()
    return data


def render_message_body(msg: ChatMessage, chat_dir: str, html_dir: str) -> Optional[str]:
    """render a single message body, including its attachments"""

    def get_attachment(match: re.Match[str]):
        if not msg.attachments:
            print("  Error: attachment HTML but no attachments in msg object")
            return "Attachment (no attachment data)<br/>"

        attachment_id = match.group(1)
        attachment = [a for a in msg.attachments if a.id == attachment_id][0]
        if attachment.content_type == "reference":
            return f"Attachment: <a href='{attachment.content_url}' data-attachment-id='{attachment.id}'>{attachment.name}</a><br/>"
        elif attachment.content_type == "messageReference" and attachment.content:
            ref = json.loads(attachment.content)
            return f"<blockquote class='message-reference' data-attachment-id='{attachment.id}'>{ref['messageSender']['user']['displayName']}: {ref['messagePreview']}</blockquote>"
        elif attachment.content_type == "application/vnd.microsoft.card.codesnippet":
            hosted_content_id = get_hosted_content_id(attachment)
            content = render_hosted_content(msg, hosted_content_id, chat_dir)
            return f"<div class='hosted-content' data-attachment-id='{attachment.id}' data-hosted-content-id='{hosted_content_id}'><pre><code>{content}</code></pre></div>"
        else:
            return f"Attachment (raw data): {pprint.pformat(attachment)}<br/>"

    def get_image(match: re.Match[str]):
        whole_match = match.group(0)
        url = match.group(1)
        if "https://graph.microsoft.com/v1.0/chats/" in url and msg.id:
            hosted_content_id = url.split("/")[-2]
            filename = get_hosted_content_filename(msg.id, hosted_content_id)
            with open(os.path.join(chat_dir, filename), "rb") as f:
                # TODO: not all images are actually png but this seems to work anyway
                data = "data:image/png;base64," + base64.b64encode(f.read()).decode(
                    "utf-8"
                )
                return (
                    whole_match.replace(url, data)
                    + f" data-hosted-content-id='{hosted_content_id}'"
                )
        else:
            return whole_match

    if msg.body and msg.body.content:
        v = msg.body.content
        if v[0:3].lower() != "<p>":
            v = f"<p>{v}</p>"

        v = re.sub('<emoji.+?alt="(.+?)".+?></emoji>', r"\g<1>", v)

        v = re.sub('<attachment id="(.+?)"></attachment>', get_attachment, v)

        # loosey-goosey matching here :(
        v = re.sub('src="(.+?)"', get_image, v)
        return v

    return None


def render_chat(chat: Chat, output_dir: str):
    """
    render a single chat to an html file. returns the name of the file rendered.
    """

    # read all the msgs for the chat, order them in chron order

    html_dir = os.path.join(output_dir, "html")
    chat_dir = os.path.join(output_dir, "data", sanitize_filename(chat.id) or "unknown_id")

    messages_files = sorted(glob.glob(os.path.join(chat_dir, f"msg_*.json")))
    msgs: list[dict[str, ChatMessage | str | None]] = []
    for path in messages_files:
        with open(path, "rb") as f:
            
            kiota_factory = kiota_serialization_json.json_parse_node_factory.JsonParseNodeFactory()
            kiota_parsenode = kiota_factory.get_root_parse_node(kiota_factory.get_valid_content_type(), f.read())
            msg = kiota_parsenode.get_object_value(ChatMessage.create_from_discriminator_value(kiota_parsenode))
            
            msgs.append(
                {"obj": msg, "content": render_message_body(msg, chat_dir, html_dir)}
            )

    # write out the html file

    filename = sanitize_filename(f"{chat.id}.html")

    path = os.path.join(html_dir, filename)

    with open(path, "w") as f:
        print(f"Writing {path}")
        template = get_jinja_env().get_template("chat.jinja")
        f.write(
            template.render(
                chat=chat,
                member_list_str=get_member_list(chat),
                messages=msgs,
            )
        )
    return filename


def render_all(output_dir: str):
    """render all the chats to html files"""

    all_chats: list[dict[str, str]] = []

    makedir(os.path.join(output_dir, "html"))

    chat_files = sorted(glob.glob(os.path.join(output_dir, "data", "*.json")))
    for path in chat_files:
        with open(path, "rb") as f:
            kiota_factory = kiota_serialization_json.json_parse_node_factory.JsonParseNodeFactory()
            kiota_parsenode = kiota_factory.get_root_parse_node(kiota_factory.get_valid_content_type(), f.read())
            chat = kiota_parsenode.get_object_value(Chat.create_from_discriminator_value(kiota_parsenode))

            filename = render_chat(chat, output_dir)

            chat_name = get_chat_name(chat)

            all_chats.append({"filename": filename, "chat_name": chat_name})

    all_chats = sorted(all_chats, key=lambda d: d["chat_name"])

    index_file = os.path.join(output_dir, "html", "index.html")

    with open(index_file, "w") as f:
        print(f"Writing {index_file}")
        template = get_jinja_env().get_template("index.jinja")
        f.write(
            template.render(
                chats=all_chats,
            )
        )


def get_graph_client() -> GraphServiceClient:
    if not client_id:
        print("Error: the CLIENT_ID environment variable isn't set")
        sys.exit(1)

    credential = InteractiveBrowserCredential(client_id=client_id)
    scopes = ["Chat.Read", "Sites.Read.All", "Files.Read.All"]
    client = GraphServiceClient(credentials=credential, scopes=scopes)
    return client


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=["download", "generate_html"])
    parser.add_argument("--output-dir", type=str, default="archive")
    parser.add_argument(
        "--force", help="download all msgs, not just 'newest' ones", action="store_true"
    )
    args = parser.parse_args()

    if args.command == "download":
        asyncio.run(download_all(args.output_dir, args.force))
    elif args.command == "generate_html":
        render_all(args.output_dir)
    else:
        print(f"Error: unrecognized command '{args.command}'")


if __name__ == "__main__":
    main()
