import os
import re
import json

import discord
from dotenv import load_dotenv


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()

TOKEN = os.getenv("DISCORD_TOKEN")

if not TOKEN:
    raise RuntimeError(
        "No se encontró DISCORD_TOKEN en el archivo .env"
    )


# ============================================================
# GOLDEN AGE MANAGER CONFIGURATION
# ============================================================

# ------------------------------------------------------------
# MENTOR SYSTEM
# ------------------------------------------------------------

MENTOR_APPLICATION_FORUM_ID = 1543919280830029854
APPROVED_MENTORS_FORUM_ID = 1543921189129293854

MENTOR_APPROVAL_MAPPING = {
    "Approved Technique": "Technique",
    "Approved Custom Form": "Custom-Form",
    "Approved Passive": "Passive-Skill",
}

MENTOR_REJECTED_TAG = "Rejected"


# ------------------------------------------------------------
# CHARACTER SYSTEM
# ------------------------------------------------------------

CHARACTER_APPLICATION_FORUM_ID = 1543968641853300827
APPROVED_CHARACTERS_FORUM_ID = 1543909774062518333

CHARACTER_ACCEPTED_TAG = "Accepted"
CHARACTER_REJECTED_TAG = "Rejected"


# Lo que puede escribir el usuario -> tag real
RACE_MAPPING = {
    # Saiyans
    "saiyan": "Saiyan",
    "pure blood saiyan": "Saiyan",
    "pure-blood saiyan": "Saiyan",
    "pureblood saiyan": "Saiyan",
    "saiyan hybrid": "Saiyan",
    "hybrid saiyan": "Saiyan",
    "half saiyan": "Saiyan",
    "half-saiyan": "Saiyan",

    # Frost Demon / Arcosian
    "frost demon": "Frost-Demon",
    "frost-demon": "Frost-Demon",
    "frost demon clan": "Frost-Demon",
    "arcosian": "Frost-Demon",
    "acrosian": "Frost-Demon",
    "frieza race": "Frost-Demon",
    "frieza clan": "Frost-Demon",

    # Namekian
    "namekian": "Namekian",
    "namek": "Namekian",
    "namekian race": "Namekian",

    # Human
    "human": "Human",
    "earthling": "Human",
    "earthling human": "Human",

    # Majin
    "majin": "Majin",

    # Bio-Android
    "bio android": "Bio-Android",
    "bio-android": "Bio-Android",
    "bioandroid": "Bio-Android",

    # Android
    "android": "Android",
}


# ============================================================
# PERSISTENT STATE / DUPLICATE PROTECTION
# ============================================================

STATE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "processed_applications.json"
)


def load_processed_applications():
    default_state = {
        "mentor": {},
        "character": {}
    }

    if not os.path.exists(STATE_FILE):
        return default_state

    try:
        with open(
            STATE_FILE,
            "r",
            encoding="utf-8"
        ) as file:
            data = json.load(file)

        if not isinstance(data, dict):
            return default_state

        data.setdefault("mentor", {})
        data.setdefault("character", {})

        return data

    except (
        json.JSONDecodeError,
        OSError
    ) as error:

        print(
            f"⚠️ No se pudo leer processed_applications.json: "
            f"{error}"
        )

        return default_state


PROCESSED_APPLICATIONS = load_processed_applications()


def save_processed_applications():

    try:
        with open(
            STATE_FILE,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                PROCESSED_APPLICATIONS,
                file,
                indent=4,
                ensure_ascii=False
            )

    except OSError as error:

        print(
            f"⚠️ No se pudo guardar processed_applications.json: "
            f"{error}"
        )


def application_was_processed(
    application_type: str,
    thread_id: int
) -> bool:

    return (
        str(thread_id)
        in PROCESSED_APPLICATIONS.get(
            application_type,
            {}
        )
    )


def mark_application_processed(
    application_type: str,
    source_thread_id: int,
    destination_thread_id: int,
    status: str = "approved"
):

    PROCESSED_APPLICATIONS.setdefault(
        application_type,
        {}
    )

    PROCESSED_APPLICATIONS[
        application_type
    ][str(source_thread_id)] = {
        "destination_thread_id": destination_thread_id,
        "status": status
    }

    save_processed_applications()


# ============================================================
# DISCORD CLIENT
# ============================================================

intents = discord.Intents.default()
intents.message_content = True

client = discord.Client(
    intents=intents
)


# ============================================================
# GENERAL UTILITIES
# ============================================================

def normalize(text: str) -> str:

    if not text:
        return ""

    text = text.strip().casefold()

    text = re.sub(
        r"[*_`~]",
        "",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def clean_value(text: str) -> str:

    if not text:
        return ""

    text = text.strip()

    text = re.sub(
        r"^[\s\-•]+",
        "",
        text
    )

    return text.strip()


def find_forum_tag(
    forum: discord.ForumChannel,
    tag_name: str
):

    wanted = normalize(tag_name)

    for tag in forum.available_tags:

        if normalize(tag.name) == wanted:
            return tag

    return None


async def get_channel(channel_id: int):

    channel = client.get_channel(
        channel_id
    )

    if channel is not None:
        return channel

    try:

        return await client.fetch_channel(
            channel_id
        )

    except discord.HTTPException as error:

        print(
            f"❌ No pude obtener el canal "
            f"{channel_id}: {error}"
        )

        return None


async def get_starter_message(
    thread: discord.Thread
):

    try:

        return await thread.fetch_message(
            thread.id
        )

    except (
        discord.NotFound,
        discord.Forbidden
    ):
        pass

    except discord.HTTPException:
        pass


    try:

        async for message in thread.history(
            limit=1,
            oldest_first=True
        ):

            return message

    except discord.HTTPException:
        pass

    return None


def is_image_attachment(
    attachment: discord.Attachment
) -> bool:

    if (
        attachment.content_type
        and attachment.content_type.startswith(
            "image/"
        )
    ):
        return True

    image_extensions = (
        ".png",
        ".jpg",
        ".jpeg",
        ".webp",
        ".gif",
    )

    return attachment.filename.lower().endswith(
        image_extensions
    )


async def prepare_attachments(
    starter_message: discord.Message
):

    files = []
    thumbnail_filename = None

    for attachment in (
        starter_message.attachments[:10]
    ):

        try:

            file = await attachment.to_file(
                use_cached=True
            )

            files.append(file)

            if (
                thumbnail_filename is None
                and is_image_attachment(
                    attachment
                )
            ):

                thumbnail_filename = (
                    file.filename
                )

            print(
                f"📎 Attachment copied: "
                f"{attachment.filename}"
            )

        except (
            discord.HTTPException,
            discord.Forbidden
        ):

            print(
                f"⚠️ Couldn't copy attachment: "
                f"{attachment.filename}"
            )

    return files, thumbnail_filename


def add_long_field(
    embed: discord.Embed,
    name: str,
    value: str,
    inline: bool = False,
    max_chunks: int = 4
):

    if not value:
        return

    value = value.strip()
    chunks = []

    while (
        value
        and len(chunks) < max_chunks
    ):

        if len(value) <= 1024:

            chunks.append(value)
            break

        cut = value.rfind(
            "\n",
            0,
            1024
        )

        if cut < 500:

            cut = value.rfind(
                " ",
                0,
                1024
            )

        if cut < 500:
            cut = 1024

        chunks.append(
            value[:cut].strip()
        )

        value = value[
            cut:
        ].strip()

    for index, chunk in enumerate(
        chunks
    ):

        field_name = (
            name
            if index == 0
            else f"{name} (continued)"
        )

        embed.add_field(
            name=field_name,
            value=chunk,
            inline=inline
        )


# ============================================================
# GENERIC LABELED-FIELD PARSER
# ============================================================

def parse_labeled_fields(
    content: str,
    aliases: dict
):

    fields = {
        key: None
        for key in aliases
    }

    patterns = []

    for (
        field_name,
        field_aliases
    ) in aliases.items():

        for alias in field_aliases:

            patterns.append(
                (
                    field_name,
                    alias
                )
            )

    patterns.sort(
        key=lambda item: len(
            item[1]
        ),
        reverse=True
    )

    alternatives = "|".join(
        f"(?P<F{i}>{pattern})"
        for i, (_, pattern)
        in enumerate(patterns)
    )

    combined = re.compile(
        rf"""
        (?<!\w)
        (?:{alternatives})
        \s*:\s*
        """,
        re.IGNORECASE | re.VERBOSE
    )

    matches = list(
        combined.finditer(content)
    )

    for index, match in enumerate(
        matches
    ):

        detected_index = None

        for group_index in range(
            len(patterns)
        ):

            if match.group(
                f"F{group_index}"
            ) is not None:

                detected_index = (
                    group_index
                )

                break

        if detected_index is None:
            continue

        field_name = patterns[
            detected_index
        ][0]

        start = match.end()

        if index + 1 < len(matches):

            end = matches[
                index + 1
            ].start()

        else:

            end = len(content)

        value = clean_value(
            content[start:end]
        )

        if value:

            fields[
                field_name
            ] = value

    return fields


# ============================================================
# MENTOR PARSER
# ============================================================

def parse_mentor_application(
    content: str
):

    aliases = {

        "character": [
            r"character\s*name",
            r"character",
        ],

        "skill": [
            r"skill\s*i\s*want\s*to\s*mentor",
            r"mentor\s*skill",
            r"skill",
        ],

        "reason": [
            r"why\s*do\s*i\s*want\s*to\s*be\s*the\s*mentor\s*of\s*this\s*skill\??",
            r"lore\s*reason\s*for\s*mentorship",
            r"reason\s*for\s*mentorship",
            r"mentor\s*reasoning",
            r"mentor\s*background",
            r"reason",
        ],

        "hours": [
            r"availability",
            r"hours",
        ],

        "lore_event": [
            r"lore\s*event",
        ],
    }

    return parse_labeled_fields(
        content,
        aliases
    )


# ============================================================
# CHARACTER PARSER
# ============================================================

def parse_character_application(
    content: str
):

    aliases = {

        "character_name": [
            r"character\s*name",
            r"character",
            r"name",
        ],

        "race": [
            r"race",
            r"species",
        ],

        "alias": [
            r"alias",
            r"nickname",
        ],

        "class": [
            r"class",
        ],

        "age": [
            r"age",
        ],

        "sex": [
            r"sex",
            r"gender",
        ],

        "family": [
            r"family",
        ],

        "alignment": [
            r"alignment",
        ],

        "fighting_style": [
            r"fighting\s*style",
            r"combat\s*style",
        ],

        "backstory": [
            r"lore\s*/\s*story\s*summary",
            r"lore\s*story\s*summary",
            r"lore\s*summary",
            r"story\s*summary",
            r"backstory\s*/\s*creation",
            r"backstory",
            r"lore",
        ],
    }

    return parse_labeled_fields(
        content,
        aliases
    )


# ============================================================
# RACE DETECTION
# ============================================================

def detect_race_tag(
    race_value: str
):

    if not race_value:
        return None

    race_normalized = normalize(
        race_value
    )

    matches = []

    for (
        written_race,
        target_tag
    ) in RACE_MAPPING.items():

        written_normalized = normalize(
            written_race
        )

        if (
            race_normalized
            == written_normalized
            or written_normalized
            in race_normalized
        ):

            matches.append(
                (
                    len(
                        written_normalized
                    ),
                    target_tag
                )
            )

    if not matches:
        return None

    matches.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return matches[0][1]


# ============================================================
# REJECTION AUTOMATION
# ============================================================

async def reject_application(
    thread: discord.Thread,
    application_label: str
):

    try:

        await thread.send(
            f"## ❌ {application_label} Rejected\n\n"
            f"Your {application_label.lower()} has been "
            "rejected by the Staff Team.\n\n"
            "If you need clarification regarding this "
            "decision, please contact a member of Staff."
        )

        await thread.edit(
            archived=True,
            locked=True,
            reason=(
                "Golden Age Manager: "
                f"{application_label} rejected"
            )
        )

        print(
            f"❌ {application_label} rejected."
        )

        print(
            "🔒 Application archived and locked."
        )

    except discord.Forbidden:

        print(
            "❌ Golden Age Manager doesn't have "
            "permission to reject this application."
        )

    except discord.HTTPException as error:

        print(
            f"❌ Discord API error while rejecting: "
            f"{error}"
        )


# ============================================================
# BOT READY
# ============================================================

@client.event
async def on_ready():

    print()
    print("=" * 65)
    print("🐉 GOLDEN AGE MANAGER")
    print("=" * 65)

    print(
        f"✅ Connected as: {client.user}"
    )

    print(
        f"🆔 Bot ID: {client.user.id}"
    )

    print()

    forum_ids = {

        "mentor-application":
            MENTOR_APPLICATION_FORUM_ID,

        "approved-mentors":
            APPROVED_MENTORS_FORUM_ID,

        "character-application":
            CHARACTER_APPLICATION_FORUM_ID,

        "approved-characters":
            APPROVED_CHARACTERS_FORUM_ID,
    }

    for (
        forum_name,
        forum_id
    ) in forum_ids.items():

        forum = await get_channel(
            forum_id
        )

        print(
            f"📁 {forum_name}"
        )

        if isinstance(
            forum,
            discord.ForumChannel
        ):

            for tag in (
                forum.available_tags
            ):

                print(
                    f"   🏷️ {tag.name}"
                )

        else:

            print(
                "   ❌ Forum not found."
            )

        print()

    print("=" * 65)

    print(
        "🤖 Waiting for applications..."
    )

    print("=" * 65)
    print()


# ============================================================
# MAIN THREAD EVENT
# ============================================================

@client.event
async def on_thread_update(
    before: discord.Thread,
    after: discord.Thread
):

    if (
        after.parent_id
        == MENTOR_APPLICATION_FORUM_ID
    ):

        await handle_mentor_application(
            before,
            after
        )

        return


    if (
        after.parent_id
        == CHARACTER_APPLICATION_FORUM_ID
    ):

        await handle_character_application(
            before,
            after
        )

        return


# ============================================================
# MENTOR AUTOMATION
# ============================================================

async def handle_mentor_application(
    before: discord.Thread,
    after: discord.Thread
):

    before_tags = {
        normalize(tag.name)
        for tag in before.applied_tags
    }

    after_tags = {
        normalize(tag.name)
        for tag in after.applied_tags
    }

    newly_added = (
        after_tags - before_tags
    )

    if not newly_added:
        return


    # --------------------------------------------------------
    # REJECTED
    # --------------------------------------------------------

    if normalize(
        MENTOR_REJECTED_TAG
    ) in newly_added:

        await reject_application(
            after,
            "Mentor Application"
        )

        return


    # --------------------------------------------------------
    # APPROVED
    # --------------------------------------------------------

    selected_source_tag = None
    destination_tag_name = None

    for (
        source_tag,
        destination_tag
    ) in MENTOR_APPROVAL_MAPPING.items():

        if normalize(
            source_tag
        ) in newly_added:

            selected_source_tag = (
                source_tag
            )

            destination_tag_name = (
                destination_tag
            )

            break

    if selected_source_tag is None:
        return


    # --------------------------------------------------------
    # DUPLICATE PROTECTION
    # --------------------------------------------------------

    if application_was_processed(
        "mentor",
        after.id
    ):

        print(
            "⚠️ This mentor application was "
            "already processed. Ignoring duplicate."
        )

        return


    print()
    print("=" * 65)
    print("🎓 MENTOR APPROVAL DETECTED")

    print(
        f"📋 {after.name}"
    )

    print(
        f"🏷️ {selected_source_tag}"
    )

    print("=" * 65)


    approved_forum = await get_channel(
        APPROVED_MENTORS_FORUM_ID
    )

    if not isinstance(
        approved_forum,
        discord.ForumChannel
    ):

        print(
            "❌ approved-mentors not found."
        )

        return


    destination_tag = find_forum_tag(
        approved_forum,
        destination_tag_name
    )

    if destination_tag is None:

        print(
            f"❌ Destination tag "
            f"'{destination_tag_name}' "
            f"not found."
        )

        return


    starter_message = (
        await get_starter_message(
            after
        )
    )

    if starter_message is None:

        print(
            "❌ Starter message not found."
        )

        return


    author = starter_message.author


    application = (
        parse_mentor_application(
            starter_message.content or ""
        )
    )


    character = application[
        "character"
    ]

    skill = application[
        "skill"
    ]

    reason = application[
        "reason"
    ]

    hours = application[
        "hours"
    ]

    lore_event = application[
        "lore_event"
    ]


    # --------------------------------------------------------
    # EMBED
    # --------------------------------------------------------

    embed = discord.Embed(
        title="✅ Approved Mentor"
    )

    embed.set_author(
        name=author.display_name,
        icon_url=author.display_avatar.url
    )


    if character:

        embed.description = (
            f"## {character}\n"
            f"**{destination_tag_name} Mentor**"
        )

    else:

        embed.description = (
            f"**{destination_tag_name} Mentor**"
        )


    if skill:

        add_long_field(
            embed,
            "🎯 Skill",
            skill
        )


    if reason:

        add_long_field(
            embed,
            "📖 Mentor Background",
            reason
        )


    if hours:

        embed.add_field(
            name="⏱️ Availability",
            value=hours[:1024],
            inline=True
        )


    embed.add_field(
        name="🎓 Mentorship",
        value=destination_tag_name,
        inline=True
    )


    if lore_event:

        add_long_field(
            embed,
            "🌎 Lore Event",
            lore_event
        )


    embed.add_field(
        name="🔗 Application",
        value=(
            f"[View original application]"
            f"({starter_message.jump_url})"
        ),
        inline=False
    )


    embed.set_footer(
        text=(
            "Dragon Ball: Golden Age • "
            "Mentorship Program"
        )
    )


    # --------------------------------------------------------
    # ATTACHMENTS
    # --------------------------------------------------------

    files, thumbnail_filename = (
        await prepare_attachments(
            starter_message
        )
    )


    if thumbnail_filename:

        embed.set_thumbnail(
            url=(
                f"attachment://"
                f"{thumbnail_filename}"
            )
        )


    # --------------------------------------------------------
    # CREATE APPROVED MENTOR
    # --------------------------------------------------------

    try:

        result = (
            await approved_forum.create_thread(
                name=after.name[:100],
                embed=embed,
                applied_tags=[
                    destination_tag
                ],
                files=files
            )
        )


        approved_thread = (
            result.thread
        )


        mark_application_processed(
            "mentor",
            after.id,
            approved_thread.id
        )


        print(
            f"✅ Mentor profile created: "
            f"{approved_thread.name}"
        )


        await after.send(
            "## ✅ Application Approved\n\n"
            "Your mentor application has "
            "been accepted.\n\n"
            f"🎓 **Mentorship Type:** "
            f"`{destination_tag_name}`\n\n"
            "Your mentor profile has been "
            f"published in "
            f"{approved_forum.mention}."
        )


        await after.edit(
            archived=True,
            locked=True,
            reason=(
                "Golden Age Manager: "
                "Mentor approved"
            )
        )


        print(
            "🔒 Original mentor application "
            "archived."
        )

        print(
            "✅ Mentor automation completed."
        )

        print()


    except discord.Forbidden:

        print(
            "❌ Permission error during "
            "mentor approval."
        )


    except discord.HTTPException as error:

        print(
            f"❌ Discord API error: {error}"
        )


    except Exception as error:

        print(
            f"❌ Unexpected mentor error: "
            f"{repr(error)}"
        )


# ============================================================
# CHARACTER AUTOMATION
# ============================================================

async def handle_character_application(
    before: discord.Thread,
    after: discord.Thread
):

    before_tags = {
        normalize(tag.name)
        for tag in before.applied_tags
    }

    after_tags = {
        normalize(tag.name)
        for tag in after.applied_tags
    }

    newly_added = (
        after_tags - before_tags
    )

    if not newly_added:
        return


    # --------------------------------------------------------
    # REJECTED
    # --------------------------------------------------------

    if normalize(
        CHARACTER_REJECTED_TAG
    ) in newly_added:

        await reject_application(
            after,
            "Character Application"
        )

        return


    # --------------------------------------------------------
    # ACCEPTED
    # --------------------------------------------------------

    accepted_normalized = normalize(
        CHARACTER_ACCEPTED_TAG
    )

    if accepted_normalized not in newly_added:
        return


    # --------------------------------------------------------
    # DUPLICATE PROTECTION
    # --------------------------------------------------------

    if application_was_processed(
        "character",
        after.id
    ):

        print(
            "⚠️ This character application was "
            "already processed. Ignoring duplicate."
        )

        return


    print()
    print("=" * 65)
    print("🐉 CHARACTER APPROVAL DETECTED")

    print(
        f"📋 {after.name}"
    )

    print("=" * 65)


    approved_forum = await get_channel(
        APPROVED_CHARACTERS_FORUM_ID
    )


    if not isinstance(
        approved_forum,
        discord.ForumChannel
    ):

        print(
            "❌ approved-characters not found."
        )

        return


    starter_message = (
        await get_starter_message(
            after
        )
    )


    if starter_message is None:

        print(
            "❌ Character starter message "
            "not found."
        )

        return


    author = starter_message.author


    content = (
        starter_message.content
        or ""
    )


    application = (
        parse_character_application(
            content
        )
    )


    character_name = (
        application["character_name"]
    )

    race = application[
        "race"
    ]

    alias = application[
        "alias"
    ]

    character_class = (
        application["class"]
    )

    age = application[
        "age"
    ]

    sex = application[
        "sex"
    ]

    family = application[
        "family"
    ]

    alignment = (
        application["alignment"]
    )

    fighting_style = (
        application["fighting_style"]
    )

    backstory = (
        application["backstory"]
    )


    # --------------------------------------------------------
    # RACE
    # --------------------------------------------------------

    race_tag_name = detect_race_tag(
        race
    )


    if race_tag_name is None:

        print(
            f"❌ Race could not be detected. "
            f"Detected value: {race!r}"
        )


        try:

            await after.send(
                "⚠️ **Golden Age Manager could "
                "not determine this character's "
                "race.**\n\n"
                "Please make sure the application "
                "contains a field such as:\n"
                "`Race: Saiyan`\n"
                "`Race: Namekian`\n"
                "`Race: Arcosian`\n"
                "`Race: Majin`\n"
                "etc.\n\n"
                "The application was **not "
                "published or archived**."
            )

        except discord.HTTPException:
            pass

        return


    destination_tag = find_forum_tag(
        approved_forum,
        race_tag_name
    )


    if destination_tag is None:

        print(
            f"❌ Race tag '{race_tag_name}' "
            f"does not exist in "
            f"approved-characters."
        )

        return


    print(
        f"🧬 Race detected: "
        f"{race} -> {race_tag_name}"
    )


    # --------------------------------------------------------
    # EMBED
    # --------------------------------------------------------

    embed = discord.Embed(
        title="✅ Approved Character"
    )


    embed.set_author(
        name=author.display_name,
        icon_url=author.display_avatar.url
    )


    display_name = (
        character_name
        or after.name
    )


    embed.description = (
        f"## {display_name}\n"
        f"**{race_tag_name}**"
    )


    if alias:

        embed.add_field(
            name="🏷️ Alias",
            value=alias[:1024],
            inline=True
        )


    embed.add_field(
        name="🧬 Race",
        value=race_tag_name,
        inline=True
    )


    if character_class:

        embed.add_field(
            name="⚔️ Class",
            value=character_class[:1024],
            inline=True
        )


    if age:

        embed.add_field(
            name="🎂 Age",
            value=age[:1024],
            inline=True
        )


    if sex:

        embed.add_field(
            name="👤 Sex",
            value=sex[:1024],
            inline=True
        )


    if alignment:

        embed.add_field(
            name="⚖️ Alignment",
            value=alignment[:1024],
            inline=True
        )


    if family:

        embed.add_field(
            name="👥 Family",
            value=family[:1024],
            inline=False
        )


    if fighting_style:

        add_long_field(
            embed,
            "🥋 Fighting Style",
            fighting_style
        )


    if backstory:

        add_long_field(
            embed,
            "📖 Lore / Backstory",
            backstory
        )


    embed.add_field(
        name="🔗 Application",
        value=(
            f"[View original application]"
            f"({starter_message.jump_url})"
        ),
        inline=False
    )


    embed.set_footer(
        text=(
            "Dragon Ball: Golden Age • "
            "Character Registry"
        )
    )


    # --------------------------------------------------------
    # ATTACHMENTS
    # --------------------------------------------------------

    files, thumbnail_filename = (
        await prepare_attachments(
            starter_message
        )
    )


    if thumbnail_filename:

        embed.set_thumbnail(
            url=(
                f"attachment://"
                f"{thumbnail_filename}"
            )
        )


    # --------------------------------------------------------
    # CREATE APPROVED CHARACTER
    # --------------------------------------------------------

    try:

        result = (
            await approved_forum.create_thread(
                name=after.name[:100],
                embed=embed,
                applied_tags=[
                    destination_tag
                ],
                files=files
            )
        )


        approved_thread = (
            result.thread
        )


        mark_application_processed(
            "character",
            after.id,
            approved_thread.id
        )


        print(
            f"✅ Approved character created: "
            f"{approved_thread.name}"
        )


        await after.send(
            "## ✅ Character Approved\n\n"
            "Your character application has "
            "been accepted.\n\n"
            f"🧬 **Race:** "
            f"`{race_tag_name}`\n\n"
            "Your character profile has been "
            f"published in "
            f"{approved_forum.mention}."
        )


        await after.edit(
            archived=True,
            locked=True,
            reason=(
                "Golden Age Manager: "
                "Character approved"
            )
        )


        print(
            "🔒 Original character application "
            "archived."
        )

        print(
            "✅ Character automation "
            "completed successfully."
        )

        print()


    except discord.Forbidden:

        print()
        print(
            "❌ CHARACTER PERMISSION ERROR"
        )

        print(
            "Golden Age Manager does not "
            "have enough permissions."
        )

        print()


    except discord.HTTPException as error:

        print()
        print(
            "❌ CHARACTER DISCORD API ERROR"
        )

        print(error)

        print()


    except Exception as error:

        print()
        print(
            "❌ CHARACTER UNEXPECTED ERROR"
        )

        print(
            repr(error)
        )

        print()


# ============================================================
# START GOLDEN AGE MANAGER
# ============================================================

client.run(TOKEN)