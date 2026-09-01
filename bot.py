import os
import re
import json
import asyncio
from datetime import datetime
from zoneinfo import ZoneInfo

import discord
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


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


# ------------------------------------------------------------
# MUTANT SYSTEM
# ------------------------------------------------------------

MUTANT_APPLICATION_FORUM_ID = 1544184281474142258
APPROVED_MUTANTS_FORUM_ID = 1544184521862291516

MUTANT_ACCEPTED_TAG = "Accepted"
MUTANT_REJECTED_TAG = "Rejected"
MUTANT_PENDING_TAG = "Pending"
MUTANT_MIGRATION_TAG = "Migration"
MUTANT_NEW_APPLICATION_TAG = "New Application"

# approved-mutants uses public/lore race labels. Android mutants are grouped
# with Human because the current Mutant Count does not have an Android row.
MUTANT_APPROVED_TAG_BY_CHARACTER_RACE = {
    "Saiyan": "Saiyan",
    "Human": "Human",
    "Android": "Human",
    "Namekian": "Namekian",
    "Frost-Demon": "Arcosian",
    "Bio-Android": "Bio-Android",
    "Majin": "Majin",
}


# ------------------------------------------------------------
# ACCESS / ROLE SYSTEM
# ------------------------------------------------------------

# Member + Verified => Community Access
COMMUNITY_ACCESS_ROLE_ID = 1544132019993190510
APPROVED_PLAYER_ROLE_ID = 1544132204370726982

MEMBER_ROLE_ID = 1543940669939261481
VERIFIED_ROLE_ID = 1543980307290791976

# Discord role assigned after a character is approved.
# Human and Android intentionally share the same Discord role.
CHARACTER_RACE_ROLE_IDS = {
    "Saiyan": 1543926682774147113,
    "Frost-Demon": 1543929886132011100,
    "Namekian": 1543929975692992513,
    "Human": 1543930046681448560,
    "Android": 1543930046681448560,
    "Majin": 1543930112037232730,
    "Bio-Android": 1543930297257693326,
}

# Cache used for Discord access restoration.
# user_id -> approved race tag (for example: "Saiyan")
# This remains member-based because Discord roles are assigned to members.
APPROVED_CHARACTER_CACHE = {}

# Source-of-truth records used by Race Count.
# source_thread_id -> {"user_id": int, "race": str, "name": str}
# IMPORTANT: Race Count is character/application-based, not unique-user-based.
APPROVED_CHARACTER_RECORDS = {}

# Legacy/member cache kept for access compatibility.
APPROVED_MUTANT_CACHE = {}

# Source-of-truth records used by Mutant Count.
# source_thread_id -> {"user_id": int, "race": str, "name": str}
APPROVED_MUTANT_RECORDS = {}

# Prevents the initial backfill from running more than once per process.
INITIAL_ACCESS_SYNC_DONE = False

# Prevents duplicate role operations when Discord emits on_member_update
# in response to a role change made by this bot.
ACCESS_SYNC_IN_PROGRESS = set()

# ------------------------------------------------------------
# RACE COUNT SYSTEM
# ------------------------------------------------------------

RACE_COUNT_CHANNEL_ID = 1543938120301551626

# Counts are based on Accepted Character Applications, not raw race roles.
# This prevents self-assigned / legacy race roles from consuming slots.
RACE_COUNT_LIMITS = {
    "Saiyan": ("Saiyan", 10),
    "Human": ("Human", 10),
    "Android": ("Human (Android)", 2),
    "Namekian": ("Namekian", 10),
    "Frost-Demon": ("Arcosian", 10),
    "Bio-Android": ("Bio-Android", 10),
    "Majin": ("Majin", 10),
}

# Mutant Count follows the categories used in #information.
# Android is intentionally not a separate mutant category.
MUTANT_COUNT_RACES = {
    "Saiyan": "Saiyan",
    "Human": "Human",
    "Namekian": "Namekian",
    "Frost-Demon": "Arcosian",
    "Bio-Android": "Bio-Android",
    "Majin": "Majin",
}

# Every Mutant category has a hard cap of 3 approved characters.
# Android Mutants are grouped under Human, matching MUTANT_COUNT_RACES.
MUTANT_LIMIT_PER_RACE = 3

RACE_COUNT_LOCK = asyncio.Lock()

# ------------------------------------------------------------
# ONBOARDING REPORT SYSTEM
# ------------------------------------------------------------

REPORT_COMMANDS = {
    "!onboardingreport",
    "!onboarding-report",
}

REPORT_TIMEZONE = "America/Santiago"
REPORTS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "reports"
)
REPORT_IN_PROGRESS = set()


# Lo que puede escribir el usuario -> tag real
#
# This mapping is intentionally generous because legacy applications use
# different spellings, punctuation, aliases and race names. The longest
# matching alias wins, so "bio android" is resolved before plain "android".
RACE_MAPPING = {
    # Saiyans / Hybrids
    "saiyan": "Saiyan",
    "saiyajin": "Saiyan",
    "saiyan race": "Saiyan",
    "pure saiyan": "Saiyan",
    "pure blood saiyan": "Saiyan",
    "pure-blood saiyan": "Saiyan",
    "pureblood saiyan": "Saiyan",
    "full blooded saiyan": "Saiyan",
    "full-blooded saiyan": "Saiyan",
    "saiyan hybrid": "Saiyan",
    "saiyan/hybrid": "Saiyan",
    "hybrid saiyan": "Saiyan",
    "hybrid": "Saiyan",
    "half saiyan": "Saiyan",
    "half-saiyan": "Saiyan",
    "half blood saiyan": "Saiyan",
    "half-blood saiyan": "Saiyan",
    "mutant saiyan": "Saiyan",
    "saiyan mutant": "Saiyan",

    # Frost Demon / Arcosian
    "frost demon": "Frost-Demon",
    "frost-demon": "Frost-Demon",
    "frost demon clan": "Frost-Demon",
    "arcosian": "Frost-Demon",
    "acrosian": "Frost-Demon",
    "arcosian frost demon": "Frost-Demon",
    "arcosian/frost demon": "Frost-Demon",
    "frost demon/arcosian": "Frost-Demon",
    "frieza race": "Frost-Demon",
    "freeza race": "Frost-Demon",
    "frieza clan": "Frost-Demon",
    "freeza clan": "Frost-Demon",
    "changeling": "Frost-Demon",
    "mutant arcosian": "Frost-Demon",
    "arcosian mutant": "Frost-Demon",
    "mutant frost demon": "Frost-Demon",
    "frost demon mutant": "Frost-Demon",

    # Namekian
    "namekian": "Namekian",
    "namek": "Namekian",
    "namekian race": "Namekian",
    "namekian warrior": "Namekian",
    "namekian dragon clan": "Namekian",
    "mutant namekian": "Namekian",
    "namekian mutant": "Namekian",

    # Humans
    "human": "Human",
    "earthling": "Human",
    "earthling human": "Human",
    "human earthling": "Human",
    "terran": "Human",
    "mutant human": "Human",
    "human mutant": "Human",

    # Androids / Human-Androids
    "android": "Android",
    "android race": "Android",
    "human android": "Android",
    "human/android": "Android",
    "human (android)": "Android",
    "android human": "Android",
    "cyborg": "Android",
    "artificial human": "Android",

    # Majins
    "majin": "Majin",
    "majin race": "Majin",
    "mutant majin": "Majin",
    "majin mutant": "Majin",

    # Bio-Androids
    "bio android": "Bio-Android",
    "bio-android": "Bio-Android",
    "bioandroid": "Bio-Android",
    "bio androids": "Bio-Android",
    "bio-androids": "Bio-Android",
    "biological android": "Bio-Android",
    "bio android race": "Bio-Android",
    "mutant bio android": "Bio-Android",
    "bio android mutant": "Bio-Android",
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
        "character": {},
        "mutant": {},
        "meta": {}
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
        data.setdefault("mutant", {})
        data.setdefault("meta", {})

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


def unmark_application_processed(
    application_type: str,
    source_thread_id: int
):
    """Remove a stale processed marker so an application can be rebuilt."""

    applications = PROCESSED_APPLICATIONS.get(
        application_type,
        {}
    )

    removed = applications.pop(
        str(source_thread_id),
        None
    )

    if removed is not None:
        save_processed_applications()

    return removed


async def processed_destination_exists(
    application_type: str,
    source_thread_id: int,
    expected_parent_id: int | None = None
) -> bool:
    """
    Checks whether the destination thread recorded in
    processed_applications.json still exists.

    This lets the bot self-heal when an approved profile was manually
    deleted and the source application is later accepted again.
    """

    record = PROCESSED_APPLICATIONS.get(
        application_type,
        {}
    ).get(str(source_thread_id))

    if not isinstance(record, dict):
        return False

    destination_thread_id = record.get(
        "destination_thread_id"
    )

    if not destination_thread_id:
        return False

    channel = client.get_channel(
        int(destination_thread_id)
    )

    if channel is None:
        try:
            channel = await client.fetch_channel(
                int(destination_thread_id)
            )
        except (
            discord.NotFound,
            discord.Forbidden,
            discord.HTTPException
        ):
            return False

    if not isinstance(channel, discord.Thread):
        return False

    if (
        expected_parent_id is not None
        and channel.parent_id != expected_parent_id
    ):
        return False

    return True


# ============================================================
# DISCORD CLIENT
# ============================================================

intents = discord.Intents.default()
intents.message_content = True

# Required to react to role changes and synchronize Community Access.
# This also requires SERVER MEMBERS INTENT in the Discord Developer Portal.
intents.members = True

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


async def get_guild_member(
    guild: discord.Guild,
    user_id: int
):
    """Gets a guild member from cache, falling back to the API."""

    member = guild.get_member(user_id)

    if member is not None:
        return member

    try:
        return await guild.fetch_member(user_id)

    except (
        discord.NotFound,
        discord.Forbidden,
        discord.HTTPException
    ):
        return None


def get_access_roles(guild: discord.Guild):
    """Returns the three roles used by the Community Access gate."""

    return (
        guild.get_role(MEMBER_ROLE_ID),
        guild.get_role(VERIFIED_ROLE_ID),
        guild.get_role(COMMUNITY_ACCESS_ROLE_ID),
    )


async def sync_community_access_for_member(
    member: discord.Member,
    *,
    reason: str = "Golden Age Manager: access synchronization"
):
    """
    Community Access exists only while the member has BOTH
    Member and Verified.

    Approved Player is also temporarily removed if the member loses
    either prerequisite. If the member becomes eligible again and has
    an Accepted character application, it is restored automatically.

    A small in-memory guard prevents duplicate role operations when
    Discord emits on_member_update because of a role change made here.
    """

    if member.bot:
        return False

    sync_key = (member.guild.id, member.id)

    if sync_key in ACCESS_SYNC_IN_PROGRESS:
        return False

    ACCESS_SYNC_IN_PROGRESS.add(sync_key)

    try:
        member_role, verified_role, community_role = get_access_roles(
            member.guild
        )

        if not all((member_role, verified_role, community_role)):
            print(
                "⚠️ Community Access sync skipped: one or more "
                "configured roles do not exist in this guild."
            )
            return False

        role_ids = {
            role.id
            for role in member.roles
        }

        should_have_access = (
            MEMBER_ROLE_ID in role_ids
            and VERIFIED_ROLE_ID in role_ids
        )

        has_access = COMMUNITY_ACCESS_ROLE_ID in role_ids
        changed = False

        try:
            if should_have_access and not has_access:
                await member.add_roles(
                    community_role,
                    reason=reason
                )

                print(
                    f"🎟️ Community Access granted to "
                    f"{member} ({member.id})."
                )
                changed = True

            elif not should_have_access and has_access:
                await member.remove_roles(
                    community_role,
                    reason=reason
                )

                print(
                    f"🔒 Community Access removed from "
                    f"{member} ({member.id})."
                )
                changed = True

            approved_player_role = member.guild.get_role(
                APPROVED_PLAYER_ROLE_ID
            )

            if approved_player_role is not None:
                has_approved_player = (
                    approved_player_role in member.roles
                )

                # No Member + Verified means no Minecraft/IP access.
                if not should_have_access and has_approved_player:
                    await member.remove_roles(
                        approved_player_role,
                        reason=(
                            "Golden Age Manager: access prerequisites lost"
                        )
                    )

                    print(
                        f"🔒 Approved Player temporarily removed from "
                        f"{member} because Member + Verified is incomplete."
                    )
                    changed = True

                # Restore Approved Player automatically for accepted characters.
                elif should_have_access:
                    approved_race = APPROVED_CHARACTER_CACHE.get(
                        member.id
                    )

                    if approved_race:
                        ok, result = await grant_character_access_roles(
                            member,
                            approved_race
                        )

                        if not ok:
                            print(
                                f"⚠️ Could not restore character access for "
                                f"{member}: {result}"
                            )

        except discord.Forbidden:
            print(
                f"❌ Could not update access roles for {member}. "
                "Check Manage Roles and role hierarchy."
            )

        except discord.HTTPException as error:
            print(
                f"❌ Discord API error while synchronizing "
                f"access for {member}: {error}"
            )

        return changed

    finally:
        ACCESS_SYNC_IN_PROGRESS.discard(sync_key)


async def rebuild_approved_character_cache():
    """
    Scans Accepted character applications, including archived forum posts,
    so existing approved players can be restored after a bot restart.
    """

    APPROVED_CHARACTER_CACHE.clear()
    APPROVED_CHARACTER_RECORDS.clear()

    forum = await get_channel(
        CHARACTER_APPLICATION_FORUM_ID
    )

    if not isinstance(forum, discord.ForumChannel):
        print(
            "⚠️ Could not rebuild approved character cache: "
            "character-application forum not found."
        )
        return

    threads = []
    seen_thread_ids = set()

    for thread in forum.threads:
        if thread.id not in seen_thread_ids:
            threads.append(thread)
            seen_thread_ids.add(thread.id)

    try:
        async for thread in forum.archived_threads(limit=None):
            if thread.id not in seen_thread_ids:
                threads.append(thread)
                seen_thread_ids.add(thread.id)

    except (discord.Forbidden, discord.HTTPException) as error:
        print(
            f"⚠️ Could not read every archived character application: "
            f"{error}"
        )

    accepted_name = normalize(
        CHARACTER_ACCEPTED_TAG
    )

    rejected_name = normalize(
        CHARACTER_REJECTED_TAG
    )

    accepted_threads_seen = 0
    accepted_threads_skipped = 0

    for thread in threads:
        tag_names = {
            normalize(tag.name)
            for tag in thread.applied_tags
        }

        if accepted_name not in tag_names:
            continue

        if rejected_name in tag_names:
            continue

        accepted_threads_seen += 1

        starter_message = await get_starter_message(
            thread
        )

        if starter_message is None:
            accepted_threads_skipped += 1
            print(
                f"⚠️ Accepted character skipped: '{thread.name}' "
                "has no readable starter message."
            )
            continue

        application = parse_character_application(
            starter_message.content or ""
        )

        race_tag_name = detect_race_tag(
            application.get("race")
        )

        if race_tag_name is None:
            accepted_threads_skipped += 1
            print(
                f"⚠️ Accepted character skipped: '{thread.name}' — "
                f"race could not be parsed from value "
                f"{application.get('race')!r}."
            )
            continue

        user_id = starter_message.author.id

        # Access restoration is member-based: one Discord member can only
        # receive one current race-role result from this cache.
        APPROVED_CHARACTER_CACHE[
            user_id
        ] = race_tag_name

        # Public population counts are CHARACTER-based. Each Accepted
        # application/thread occupies one race slot, even if the same Discord
        # member owns more than one approved character.
        character_display_name = (
            application.get("character_name")
            or thread.name
            or starter_message.author.display_name
        )

        APPROVED_CHARACTER_RECORDS[thread.id] = {
            "user_id": user_id,
            "race": race_tag_name,
            "name": character_display_name,
            "thread_name": thread.name,
        }

    print(
        f"🎮 Approved character access cache rebuilt: "
        f"{len(APPROVED_CHARACTER_CACHE)} unique player(s)."
    )

    print(
        f"📚 Accepted character applications counted: "
        f"{len(APPROVED_CHARACTER_RECORDS)} character(s) "
        f"from {accepted_threads_seen} Accepted thread(s)."
    )

    if accepted_threads_skipped:
        print(
            f"⚠️ Accepted applications skipped during counting: "
            f"{accepted_threads_skipped}. See warnings above."
        )

    # Per-race diagnostics make population mismatches easy to spot in Kinetic.
    diagnostic_counts = {}
    for record in APPROVED_CHARACTER_RECORDS.values():
        race_name = record.get("race", "Unknown")
        diagnostic_counts[race_name] = diagnostic_counts.get(race_name, 0) + 1

    if diagnostic_counts:
        diagnostic_text = ", ".join(
            f"{race}: {count}"
            for race, count in sorted(diagnostic_counts.items())
        )
        print(f"📊 Accepted race breakdown: {diagnostic_text}")

    await audit_approved_character_registry()


async def collect_forum_threads(forum: discord.ForumChannel):
    """Returns active + archived forum threads without duplicates."""

    threads = []
    seen_ids = set()

    for thread in forum.threads:
        if thread.id not in seen_ids:
            threads.append(thread)
            seen_ids.add(thread.id)

    try:
        async for thread in forum.archived_threads(limit=None):
            if thread.id not in seen_ids:
                threads.append(thread)
                seen_ids.add(thread.id)
    except (discord.Forbidden, discord.HTTPException) as error:
        print(
            f"⚠️ Could not read every archived thread in "
            f"#{forum.name}: {error}"
        )

    return threads


def parse_mutant_application(content: str):
    """Parses the dedicated Mutant Application template."""

    aliases = {
        "discord_username": [
            r"discord\s*username",
            r"discord",
        ],
        "character_name": [
            r"character\s*name",
            r"birth\s*name",
            r"character",
        ],
        "character_application": [
            r"character\s*application(?:\s*link)?",
            r"approved\s*character\s*application",
            r"character\s*app",
        ],
        "application_type": [
            r"application\s*type",
            r"type",
        ],
        "race": [
            r"race\s*/\s*species",
            r"species\s*/\s*race",
            r"race",
            r"species",
        ],
        "mutation_name": [
            r"mutation\s*name\s*/\s*type",
            r"mutation\s*name",
            r"mutation\s*type",
            r"mutant\s*type",
        ],
        "origin": [
            r"origin\s*of\s*the\s*mutation",
            r"mutation\s*origin",
            r"origin",
        ],
        "physical_changes": [
            r"physical\s*changes",
            r"physical\s*effects",
        ],
        "abilities": [
            r"abilities\s*/\s*advantages",
            r"advantages\s*/\s*abilities",
            r"abilities",
            r"advantages",
        ],
        "weaknesses": [
            r"weaknesses\s*/\s*limitations",
            r"limitations\s*/\s*weaknesses",
            r"weaknesses",
            r"limitations",
            r"drawbacks",
        ],
        "mutation_backstory": [
            r"mutation\s*backstory",
            r"mutation\s*lore",
            r"backstory",
        ],
        "character_effect": [
            r"how\s*does\s*the\s*mutation\s*affect\s*your\s*character\??",
            r"effect\s*on\s*character",
            r"character\s*impact",
        ],
        "previously_approved": [
            r"were\s*you\s*already\s*an\s*approved\s*mutant\s*on\s*the\s*previous\s*server\??",
            r"previously\s*approved\s*mutant",
        ],
        "previous_character_name": [
            r"previous\s*character\s*name",
        ],
        "previous_mutant_race": [
            r"previous\s*mutant\s*race",
        ],
        "proof": [
            r"proof\s*/\s*evidence",
            r"evidence\s*/\s*proof",
            r"proof",
            r"evidence",
        ],
        "additional_info": [
            r"additional\s*information",
            r"additional\s*info",
            r"anything\s*else",
        ],
    }

    return parse_labeled_fields(content, aliases)


def extract_discord_thread_id(value: str | None):
    """Extracts the channel/thread ID from a Discord message/thread URL."""

    if not value:
        return None

    # https://discord.com/channels/GUILD_ID/THREAD_ID/MESSAGE_ID
    match = re.search(
        r"(?:discord(?:app)?\.com/channels/|discord\.com/channels/)"
        r"\d+/(\d+)(?:/\d+)?",
        value,
        re.IGNORECASE,
    )
    if match:
        return int(match.group(1))

    # If the dedicated field contains only an ID, accept it as a fallback.
    match = re.fullmatch(r"\s*(\d{17,20})\s*", value)
    if match:
        return int(match.group(1))

    return None


def mutant_count_race(character_race: str | None):
    """Maps a Character race to the current Mutant Count categories."""

    if character_race == "Android":
        return "Human"

    if character_race in MUTANT_COUNT_RACES:
        return character_race

    return None


async def get_character_application_record(thread: discord.Thread):
    """Validates and parses one Accepted Character Application thread."""

    if not isinstance(thread, discord.Thread):
        return None, "The linked Character Application is not a forum thread."

    if thread.parent_id != CHARACTER_APPLICATION_FORUM_ID:
        return None, "The link does not point to #character-application."

    tag_names = {normalize(tag.name) for tag in thread.applied_tags}

    if normalize(CHARACTER_ACCEPTED_TAG) not in tag_names:
        return None, "The linked Character Application is not currently Accepted."

    if normalize(CHARACTER_REJECTED_TAG) in tag_names:
        return None, "The linked Character Application is Rejected."

    starter = await get_starter_message(thread)
    if starter is None:
        return None, "The linked Character Application could not be read."

    application = parse_character_application(starter.content or "")
    race = detect_race_tag(application.get("race"))

    if race is None:
        return None, "The linked Character Application has an unreadable race."

    return {
        "thread": thread,
        "starter": starter,
        "application": application,
        "race": race,
        "name": (
            application.get("character_name")
            or thread.name
            or starter.author.display_name
        ),
    }, None


async def resolve_character_application_for_mutant(
    mutant_author_id: int,
    mutant_application: dict,
):
    """
    Resolves the approved Character Application tied to a Mutant Application.

    Priority:
      1. Explicit Character Application link/ID in the Mutant form.
      2. Accepted applications by the same Discord author.
      3. If multiple exist, match Character Name.
    """

    explicit_value = mutant_application.get("character_application")
    explicit_thread_id = extract_discord_thread_id(explicit_value)

    if explicit_thread_id:
        channel = await get_channel(explicit_thread_id)

        if not isinstance(channel, discord.Thread):
            return None, "The Character Application link could not be resolved."

        record, error = await get_character_application_record(channel)
        if error:
            return None, error

        if record["starter"].author.id != mutant_author_id:
            return None, (
                "The linked Character Application belongs to a different "
                "Discord user."
            )

        return record, None

    forum = await get_channel(CHARACTER_APPLICATION_FORUM_ID)
    if not isinstance(forum, discord.ForumChannel):
        return None, "#character-application could not be accessed."

    matches = []
    for thread in await collect_forum_threads(forum):
        record, error = await get_character_application_record(thread)
        if error or record is None:
            continue

        if record["starter"].author.id == mutant_author_id:
            matches.append(record)

    if not matches:
        return None, (
            "No Accepted Character Application belonging to this applicant "
            "could be found."
        )

    if len(matches) == 1:
        return matches[0], None

    wanted_name = normalize(mutant_application.get("character_name") or "")
    if wanted_name:
        name_matches = [
            record for record in matches
            if normalize(record.get("name") or "") == wanted_name
            or wanted_name in normalize(record.get("name") or "")
            or normalize(record.get("name") or "") in wanted_name
        ]

        if len(name_matches) == 1:
            return name_matches[0], None

    return None, (
        "Multiple Accepted Character Applications were found for this user. "
        "Please paste the exact Character Application link in the Mutant form."
    )


async def rebuild_approved_mutant_cache():
    """
    Rebuilds Mutant Count exclusively from Accepted Mutant Applications.

    Character Applications no longer grant Mutant status by themselves.
    """

    APPROVED_MUTANT_CACHE.clear()
    APPROVED_MUTANT_RECORDS.clear()

    forum = await get_channel(MUTANT_APPLICATION_FORUM_ID)

    if not isinstance(forum, discord.ForumChannel):
        print(
            "⚠️ Could not rebuild approved mutant cache: "
            "mutant-application forum not found."
        )
        return

    accepted_name = normalize(MUTANT_ACCEPTED_TAG)
    rejected_name = normalize(MUTANT_REJECTED_TAG)
    accepted_seen = 0
    skipped = 0

    for thread in await collect_forum_threads(forum):
        tag_names = {normalize(tag.name) for tag in thread.applied_tags}

        if accepted_name not in tag_names or rejected_name in tag_names:
            continue

        accepted_seen += 1

        starter = await get_starter_message(thread)
        if starter is None:
            skipped += 1
            print(
                f"⚠️ Accepted mutant skipped: '{thread.name}' has no "
                "readable starter message."
            )
            continue

        mutant_application = parse_mutant_application(starter.content or "")
        character_record, error = await resolve_character_application_for_mutant(
            starter.author.id,
            mutant_application,
        )

        if error or character_record is None:
            skipped += 1
            print(
                f"⚠️ Accepted mutant skipped: '{thread.name}' — {error}"
            )
            continue

        race = mutant_count_race(character_record["race"])
        if race is None:
            skipped += 1
            print(
                f"⚠️ Accepted mutant skipped: '{thread.name}' — "
                f"race {character_record['race']!r} has no Mutant Count row."
            )
            continue

        name = (
            clean_value(mutant_application.get("character_name") or "")
            or clean_value(character_record.get("name") or "")
            or thread.name
        )
        name = re.sub(r"[*_`~]", "", name).splitlines()[0].strip()[:40]

        record = {
            "user_id": starter.author.id,
            "race": race,
            "character_race": character_record["race"],
            "name": name,
            "source_thread_id": thread.id,
            "character_thread_id": character_record["thread"].id,
        }

        APPROVED_MUTANT_CACHE[starter.author.id] = record
        APPROVED_MUTANT_RECORDS[thread.id] = record

    print(
        f"🧬 Approved Mutant Applications counted: "
        f"{len(APPROVED_MUTANT_RECORDS)} mutant character(s) "
        f"from {accepted_seen} Accepted thread(s)."
    )

    if skipped:
        print(
            f"⚠️ Accepted Mutant Applications skipped: {skipped}. "
            "See warnings above."
        )

    breakdown = {}
    for record in APPROVED_MUTANT_RECORDS.values():
        breakdown[record["race"]] = breakdown.get(record["race"], 0) + 1

    if breakdown:
        diagnostic = ", ".join(
            f"{MUTANT_COUNT_RACES.get(race, race)}: {count}"
            for race, count in sorted(breakdown.items())
        )
        print(f"🧬 Accepted mutant breakdown: {diagnostic}")


async def audit_approved_character_registry():
    """Compare Accepted source apps with bot-managed approved profiles."""
    forum = await get_channel(APPROVED_CHARACTERS_FORUM_ID)
    if not isinstance(forum, discord.ForumChannel):
        print("⚠️ Race audit skipped: approved-characters forum not found.")
        return

    threads = []
    seen_ids = set()
    for thread in forum.threads:
        if thread.id not in seen_ids:
            threads.append(thread)
            seen_ids.add(thread.id)

    try:
        async for thread in forum.archived_threads(limit=None):
            if thread.id not in seen_ids:
                threads.append(thread)
                seen_ids.add(thread.id)
    except (discord.Forbidden, discord.HTTPException) as error:
        print(f"⚠️ Race audit could not read all archived approved profiles: {error}")

    accepted_by_race = {race: [] for race in RACE_COUNT_LIMITS}
    for record in APPROVED_CHARACTER_RECORDS.values():
        race = record.get("race")
        name = record.get("thread_name") or record.get("name") or "Unknown"
        if race in accepted_by_race:
            accepted_by_race[race].append(name)

    registry_by_race = {race: [] for race in RACE_COUNT_LIMITS}
    for thread in threads:
        race_tag = None
        for tag in thread.applied_tags:
            for canonical_race in RACE_COUNT_LIMITS:
                if normalize(tag.name) == normalize(canonical_race):
                    race_tag = canonical_race
                    break
            if race_tag:
                break

        if race_tag is None:
            continue

        starter = await get_starter_message(thread)
        if starter is None:
            continue

        # Count only bot-created approved profiles; manual/legacy posts such as
        # Sarkala remain excluded from this comparison.
        if client.user is None or starter.author.id != client.user.id:
            continue

        registry_by_race[race_tag].append(thread.name)

    print("🔎 Race registry audit:")
    mismatch_found = False

    for race in RACE_COUNT_LIMITS:
        accepted_names = sorted(accepted_by_race[race], key=str.casefold)
        registry_names = sorted(registry_by_race[race], key=str.casefold)
        accepted_keys = {normalize(name): name for name in accepted_names}
        registry_keys = {normalize(name): name for name in registry_names}

        if len(accepted_names) != len(registry_names):
            mismatch_found = True
            display = RACE_COUNT_LIMITS[race][0]
            print(
                f"   ⚠️ {display}: Accepted apps={len(accepted_names)}, "
                f"bot-managed approved profiles={len(registry_names)}"
            )

        for key, name in registry_keys.items():
            if key not in accepted_keys:
                mismatch_found = True
                print(
                    f"      🟠 Approved profile without current Accepted source: "
                    f"[{RACE_COUNT_LIMITS[race][0]}] {name}"
                )

        for key, name in accepted_keys.items():
            if key not in registry_keys:
                mismatch_found = True
                print(
                    f"      🔴 Accepted application missing approved profile: "
                    f"[{RACE_COUNT_LIMITS[race][0]}] {name}"
                )

    if not mismatch_found:
        print("   ✅ Accepted applications and bot-managed approved profiles match.")


def get_race_counts():
    """Returns current approved-player counts by application race tag."""

    counts = {
        race_tag: 0
        for race_tag in RACE_COUNT_LIMITS
    }

    for record in APPROVED_CHARACTER_RECORDS.values():
        race_tag_name = record.get("race")
        if race_tag_name in counts:
            counts[race_tag_name] += 1

    return counts


def get_mutant_counts():
    """Returns Mutant characters grouped by canonical race."""

    counts = {
        race_tag: []
        for race_tag in MUTANT_COUNT_RACES
    }

    for record in APPROVED_MUTANT_RECORDS.values():
        race_tag = record.get("race")
        name = record.get("name") or "Unknown"

        if race_tag in counts:
            counts[race_tag].append(name)

    for names in counts.values():
        names.sort(key=str.casefold)

    return counts


def build_race_count_message():
    """Builds the public Race Count + Mutant Count from Accepted apps."""

    counts = get_race_counts()
    mutant_counts = get_mutant_counts()

    lines = [
        "## Race Count",
        "",
    ]

    for race_tag, (display_name, limit) in RACE_COUNT_LIMITS.items():
        count = counts.get(race_tag, 0)

        if count > limit:
            status = " <- OVER CAP"
        elif count >= limit:
            status = " <- FULL"
        elif count >= max(1, int(limit * 0.8)):
            status = " <- LIMITED"
        else:
            status = ""

        lines.append(
            f"**{display_name}:** {count}/{limit}{status}"
        )

    lines.extend([
        "",
        "## Mutant Count",
        "",
    ])

    for race_tag, display_name in MUTANT_COUNT_RACES.items():
        names = mutant_counts.get(race_tag, [])
        count = len(names)
        limit = MUTANT_LIMIT_PER_RACE

        if count > limit:
            status = " <- OVER CAP"
        elif count >= limit:
            status = " <- FULL"
        elif count >= 2:
            status = " <- LIMITED"
        else:
            status = ""

        if names:
            lines.append(
                f"**{display_name}:** {count}/{limit}{status} "
                f"({', '.join(names)})"
            )
        else:
            lines.append(
                f"**{display_name}:** 0/{limit}"
            )

    lines.extend([
        "",
        "-# Automatically updated by Golden Age Manager.",
        "-# Race Count uses Accepted Character Applications only.",
        "-# Mutant Count uses Accepted Mutant Applications only (max 3 per race).",
    ])

    return "\n".join(lines)


async def get_race_count_message(channel: discord.TextChannel):
    """Finds the bot-owned managed Race Count message, if one exists."""

    meta = PROCESSED_APPLICATIONS.setdefault("meta", {})
    stored_message_id = meta.get("race_count_message_id")

    if stored_message_id:
        try:
            message = await channel.fetch_message(
                int(stored_message_id)
            )

            if message.author.id == client.user.id:
                return message

        except (
            discord.NotFound,
            discord.Forbidden,
            discord.HTTPException,
            ValueError,
            TypeError,
        ):
            pass

    try:
        async for message in channel.history(limit=100):
            if (
                message.author.id == client.user.id
                and "## Race Count" in (message.content or "")
                and "Automatically updated by Golden Age Manager"
                in (message.content or "")
            ):
                meta["race_count_message_id"] = message.id
                save_processed_applications()
                return message

    except (discord.Forbidden, discord.HTTPException):
        pass

    return None


async def update_race_count_message(reason: str = "automatic update"):
    """Creates or edits Golden Age Manager's managed Race Count message."""

    async with RACE_COUNT_LOCK:
        channel = await get_channel(RACE_COUNT_CHANNEL_ID)

        if not isinstance(channel, discord.TextChannel):
            print(
                "⚠️ Race Count channel was not found or is not a text channel."
            )
            return False

        content = build_race_count_message()
        message = await get_race_count_message(channel)

        try:
            if message is None:
                message = await channel.send(content)
                PROCESSED_APPLICATIONS.setdefault(
                    "meta", {}
                )["race_count_message_id"] = message.id
                save_processed_applications()

                print(
                    f"📊 Race Count message created in #{channel.name}."
                )
            elif message.content != content:
                await message.edit(content=content)
                print(
                    f"📊 Race Count updated ({reason})."
                )
            else:
                print(
                    f"📊 Race Count already current ({reason})."
                )

            return True

        except discord.Forbidden:
            print(
                "❌ Golden Age Manager cannot create/edit the Race Count. "
                "Check View Channel, Send Messages and Read Message History."
            )
            return False

        except discord.HTTPException as error:
            print(
                f"❌ Discord API error while updating Race Count: {error}"
            )
            return False


async def sync_all_community_access():
    """Backfills Community Access for members already in the server."""

    print("🎟️ Synchronizing Community Access...")

    changed = 0

    for guild in client.guilds:
        member_role, verified_role, community_role = get_access_roles(
            guild
        )

        # Only operate in a guild that contains our configured roles.
        if not all((member_role, verified_role, community_role)):
            continue

        for member in guild.members:
            if await sync_community_access_for_member(
                member,
                reason=(
                    "Golden Age Manager: initial Community Access sync"
                )
            ):
                changed += 1

    print(
        f"✅ Community Access synchronization complete. "
        f"Changes: {changed}"
    )


def resolve_character_access_roles(
    guild: discord.Guild,
    race_tag_name: str
):
    """Resolves and validates Approved Player + detected race role."""

    approved_player_role = guild.get_role(
        APPROVED_PLAYER_ROLE_ID
    )

    race_role_id = CHARACTER_RACE_ROLE_IDS.get(
        race_tag_name
    )

    race_role = (
        guild.get_role(race_role_id)
        if race_role_id
        else None
    )

    if approved_player_role is None:
        return None, None, "Approved Player role was not found."

    if race_role_id is None:
        return (
            None,
            None,
            f"No Discord race role is configured for '{race_tag_name}'."
        )

    if race_role is None:
        return (
            None,
            None,
            f"Discord race role {race_role_id} was not found."
        )

    bot_member = guild.me

    if bot_member is None:
        return None, None, "Could not resolve the bot member in this guild."

    if not bot_member.guild_permissions.manage_roles:
        return None, None, "Golden Age Manager lacks Manage Roles."

    for target_role in (approved_player_role, race_role):
        if target_role >= bot_member.top_role:
            return (
                None,
                None,
                f"Role '{target_role.name}' is not below Golden Age "
                "Manager in the role hierarchy."
            )

    return approved_player_role, race_role, None


async def grant_character_access_roles(
    member: discord.Member,
    race_tag_name: str
):
    """
    Grants Approved Player plus the Discord role matching the approved race.
    Existing race roles are not removed in this phase.
    """

    approved_player_role, race_role, error = (
        resolve_character_access_roles(
            member.guild,
            race_tag_name
        )
    )

    if error:
        return False, error

    member_role_ids = {
        role.id
        for role in member.roles
    }

    if not (
        MEMBER_ROLE_ID in member_role_ids
        and VERIFIED_ROLE_ID in member_role_ids
    ):
        return (
            False,
            "Member + Verified are required before Approved Player "
            "can be granted."
        )

    roles_to_add = []

    if approved_player_role not in member.roles:
        roles_to_add.append(approved_player_role)

    if race_role not in member.roles:
        roles_to_add.append(race_role)

    if not roles_to_add:
        return (
            True,
            f"Already has {approved_player_role.name} and {race_role.name}."
        )

    try:
        await member.add_roles(
            *roles_to_add,
            reason=(
                "Golden Age Manager: character application approved"
            )
        )

        print(
            f"🎮 Character access granted to {member}: "
            f"{approved_player_role.name} + {race_role.name}"
        )

        return (
            True,
            f"{approved_player_role.name} + {race_role.name}"
        )

    except discord.Forbidden:
        return (
            False,
            "Discord denied role assignment. Check Manage Roles and "
            "the role hierarchy."
        )

    except discord.HTTPException as error:
        return False, f"Discord API error while assigning roles: {error}"


async def sync_processed_character_access():
    """
    Retries role delivery for character applications that this bot has
    previously recorded in processed_applications.json.
    """

    processed_characters = PROCESSED_APPLICATIONS.get(
        "character",
        {}
    )

    if not processed_characters:
        return

    print("🎮 Synchronizing processed character access...")

    successful = 0

    for source_thread_id, record in processed_characters.items():
        if record.get("status", "approved") != "approved":
            continue

        try:
            source_thread = await get_channel(
                int(source_thread_id)
            )

            if not isinstance(source_thread, discord.Thread):
                continue

            starter_message = await get_starter_message(
                source_thread
            )

            if starter_message is None:
                continue

            application = parse_character_application(
                starter_message.content or ""
            )

            race_tag_name = detect_race_tag(
                application.get("race")
            )

            if race_tag_name is None:
                continue

            member = await get_guild_member(
                source_thread.guild,
                starter_message.author.id
            )

            if member is None:
                continue

            ok, _ = await grant_character_access_roles(
                member,
                race_tag_name
            )

            if ok:
                successful += 1

        except Exception as error:
            print(
                f"⚠️ Could not backfill character application "
                f"{source_thread_id}: {error!r}"
            )

    print(
        f"✅ Processed character access sync complete. "
        f"Checked/granted: {successful}"
    )


# ============================================================
# ONBOARDING REPORT
# ============================================================

def get_application_status(thread: discord.Thread) -> str:
    """Returns a simple status derived from the forum tags."""

    tag_names = {
        normalize(tag.name)
        for tag in thread.applied_tags
    }

    if normalize(CHARACTER_ACCEPTED_TAG) in tag_names:
        return "Accepted"

    if normalize(CHARACTER_REJECTED_TAG) in tag_names:
        return "Rejected"

    if "pending" in tag_names:
        return "Pending"

    return "No Status"


async def collect_character_application_index(
    guild: discord.Guild
):
    """
    Builds user_id -> list of character applications from the source forum,
    including archived applications.
    """

    forum = await get_channel(
        CHARACTER_APPLICATION_FORUM_ID
    )

    if not isinstance(forum, discord.ForumChannel):
        raise RuntimeError(
            "character-application forum could not be found."
        )

    threads = []
    seen_thread_ids = set()

    for thread in forum.threads:
        if thread.id not in seen_thread_ids:
            threads.append(thread)
            seen_thread_ids.add(thread.id)

    try:
        async for thread in forum.archived_threads(limit=None):
            if thread.id not in seen_thread_ids:
                threads.append(thread)
                seen_thread_ids.add(thread.id)

    except (discord.Forbidden, discord.HTTPException) as error:
        print(
            f"⚠️ Could not read every archived character application "
            f"for report: {error}"
        )

    application_index = {}

    for thread in threads:
        owner_id = thread.owner_id

        if owner_id is None:
            starter_message = await get_starter_message(thread)

            if starter_message is None:
                continue

            owner_id = starter_message.author.id

        record = {
            "thread_id": thread.id,
            "thread_name": thread.name,
            "status": get_application_status(thread),
            "created_at": thread.created_at,
            "url": (
                f"https://discord.com/channels/"
                f"{guild.id}/{thread.id}"
            ),
        }

        application_index.setdefault(
            owner_id,
            []
        ).append(record)

    for records in application_index.values():
        records.sort(
            key=lambda item: item["created_at"],
            reverse=True
        )

    return application_index


async def get_complete_member_list(guild: discord.Guild):
    """Fetches a complete member list when possible, with cache fallback."""

    try:
        members = [
            member
            async for member in guild.fetch_members(limit=None)
        ]
        return members

    except (discord.Forbidden, discord.HTTPException) as error:
        print(
            f"⚠️ Full member fetch failed; using Discord cache instead: "
            f"{error}"
        )
        return list(guild.members)


def get_onboarding_anomalies(row):
    """Returns audit flags for onboarding states that deserve review."""

    anomalies = []

    if (
        row["Accepted Application"] == "Yes"
        and row["Approved Player"] == "No"
    ):
        anomalies.append("Accepted + no Approved Player")

    if (
        row["Accepted Application"] == "Yes"
        and row["Verified"] == "No"
    ):
        anomalies.append("Accepted + not Verified")

    if (
        row["Race Role(s)"]
        and row["Character Application"] == "No"
    ):
        anomalies.append("Race role + no Character Application")

    if (
        row["Character Application"] == "Yes"
        and row["Latest Application Status"] == "No Status"
    ):
        anomalies.append("Application + No Status")

    if (
        row["Approved Player"] == "Yes"
        and row["Accepted Application"] == "No"
    ):
        anomalies.append("Approved Player + no Accepted Application")

    if (
        row["Community Access"] == "Yes"
        and (
            row["Member Role"] == "No"
            or row["Verified"] == "No"
        )
    ):
        anomalies.append("Community Access without prerequisites")

    if (
        row["Approved Player"] == "Yes"
        and row["Community Access"] == "No"
    ):
        anomalies.append("Approved Player without Community Access")

    return anomalies


def get_anomaly_severity(anomalies):
    """Returns the highest severity among a member's anomaly flags."""

    if not anomalies:
        return "Clean"

    critical_flags = {
        "Approved Player + no Accepted Application",
        "Community Access without prerequisites",
        "Approved Player without Community Access",
    }

    high_flags = {
        "Accepted + no Approved Player",
    }

    if any(flag in critical_flags for flag in anomalies):
        return "Critical"

    if any(flag in high_flags for flag in anomalies):
        return "High"

    return "Review"


def build_onboarding_rows(
    guild: discord.Guild,
    members,
    application_index
):
    """Builds plain dictionaries that can safely be written to Excel."""

    race_role_ids = set(
        CHARACTER_RACE_ROLE_IDS.values()
    )

    rows = []

    for member in members:
        if member.bot:
            continue

        role_ids = {
            role.id
            for role in member.roles
        }

        has_member = MEMBER_ROLE_ID in role_ids
        is_verified = VERIFIED_ROLE_ID in role_ids
        has_community_access = (
            COMMUNITY_ACCESS_ROLE_ID in role_ids
        )
        has_approved_player = (
            APPROVED_PLAYER_ROLE_ID in role_ids
        )

        applications = application_index.get(
            member.id,
            []
        )

        has_application = bool(applications)
        application_count = len(applications)
        latest_application = (
            applications[0]
            if applications
            else None
        )

        latest_status = (
            latest_application["status"]
            if latest_application
            else "No Application"
        )

        latest_application_name = (
            latest_application["thread_name"]
            if latest_application
            else ""
        )

        latest_application_url = (
            latest_application["url"]
            if latest_application
            else ""
        )

        accepted_application = any(
            item["status"] == "Accepted"
            for item in applications
        )

        pending_application = any(
            item["status"] == "Pending"
            for item in applications
        )

        rejected_application = any(
            item["status"] == "Rejected"
            for item in applications
        )

        race_roles = [
            role.name
            for role in member.roles
            if role.id in race_role_ids
        ]

        missing_actions = []

        if not is_verified:
            missing_actions.append("Complete verification")

        if not has_application:
            missing_actions.append("Submit character application")

        if (
            accepted_application
            and is_verified
            and not has_approved_player
        ):
            missing_actions.append(
                "Approved character exists but Approved Player is missing"
            )

        if not missing_actions:
            if pending_application and not accepted_application:
                missing_actions.append("Application awaiting Staff review")
            elif rejected_application and not accepted_application:
                missing_actions.append(
                    "Character application rejected; review/resubmit if needed"
                )
            else:
                missing_actions.append("Complete")

        joined_at = ""
        if member.joined_at is not None:
            joined_at = member.joined_at.astimezone(
                ZoneInfo(REPORT_TIMEZONE)
            ).strftime("%Y-%m-%d %H:%M:%S")

        row = {
            "Discord Username": str(member),
            "Display Name": member.display_name,
            "User ID": str(member.id),
            "Joined Server": joined_at,
            "Member Role": "Yes" if has_member else "No",
            "Verified": "Yes" if is_verified else "No",
            "Community Access": (
                "Yes" if has_community_access else "No"
            ),
            "Character Application": (
                "Yes" if has_application else "No"
            ),
            "Latest Application Status": latest_status,
            "Application Count": application_count,
            "Latest Application": latest_application_name,
            "Latest Application Link": latest_application_url,
            "Accepted Application": (
                "Yes" if accepted_application else "No"
            ),
            "Approved Player": (
                "Yes" if has_approved_player else "No"
            ),
            "Race Role(s)": ", ".join(race_roles),
            "Needs Action": "; ".join(missing_actions),
            "Missing Verification + No Application": (
                "Yes"
                if (not is_verified and not has_application)
                else "No"
            ),
        }

        anomalies = get_onboarding_anomalies(row)

        row["Anomaly Count"] = len(anomalies)
        row["Anomaly Severity"] = get_anomaly_severity(anomalies)
        row["Anomaly Type(s)"] = "; ".join(anomalies)

        rows.append(row)

    rows.sort(
        key=lambda item: (
            item["Display Name"].casefold(),
            item["Discord Username"].casefold(),
        )
    )

    return rows


def style_report_sheet(
    worksheet,
    *,
    table_name: str,
    add_table: bool = True
):
    """Applies a consistent, readable Excel style."""

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )
    header_font = Font(
        color="FFFFFF",
        bold=True
    )
    thin_gray = Side(
        style="thin",
        color="D9E2F3"
    )

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = Border(
            bottom=thin_gray
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        column_letter = get_column_letter(
            column_cells[0].column
        )

        max_length = 0

        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(
                max_length,
                len(value)
            )
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 2, 10),
            42
        )

    if add_table and worksheet.max_row >= 2:
        table = Table(
            displayName=table_name,
            ref=(
                f"A1:{get_column_letter(worksheet.max_column)}"
                f"{worksheet.max_row}"
            )
        )

        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        table.tableStyleInfo = style
        worksheet.add_table(table)


def write_data_sheet(
    workbook: Workbook,
    sheet_name: str,
    rows,
    headers,
    table_name: str
):
    worksheet = workbook.create_sheet(sheet_name)
    worksheet.append(headers)

    for row in rows:
        worksheet.append([
            row.get(header, "")
            for header in headers
        ])

    style_report_sheet(
        worksheet,
        table_name=table_name,
        add_table=True
    )

    # Make links clickable in the dedicated URL column.
    if "Latest Application Link" in headers:
        link_column = headers.index(
            "Latest Application Link"
        ) + 1

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):
            cell = worksheet.cell(
                row=row_number,
                column=link_column
            )

            if cell.value:
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"

    return worksheet


def create_onboarding_excel(
    guild_name: str,
    rows,
    output_path: str,
    generated_at_text: str
):
    """Creates the final .xlsx report from plain Python data."""

    os.makedirs(
        os.path.dirname(output_path),
        exist_ok=True
    )

    workbook = Workbook()

    summary = workbook.active
    summary.title = "Summary"

    total_members = len(rows)
    verified_count = sum(
        row["Verified"] == "Yes"
        for row in rows
    )
    not_verified_count = total_members - verified_count
    has_application_count = sum(
        row["Character Application"] == "Yes"
        for row in rows
    )
    no_application_count = (
        total_members - has_application_count
    )
    both_missing_count = sum(
        row["Missing Verification + No Application"] == "Yes"
        for row in rows
    )
    incomplete_count = sum(
        (
            row["Verified"] == "No"
            or row["Character Application"] == "No"
        )
        for row in rows
    )
    accepted_count = sum(
        row["Accepted Application"] == "Yes"
        for row in rows
    )
    approved_player_count = sum(
        row["Approved Player"] == "Yes"
        for row in rows
    )
    anomaly_member_count = sum(
        row.get("Anomaly Count", 0) > 0
        for row in rows
    )
    anomaly_flag_count = sum(
        row.get("Anomaly Count", 0)
        for row in rows
    )

    summary.merge_cells("A1:D1")
    summary["A1"] = "Dragon Ball: Golden Age — Onboarding Report"
    summary["A1"].font = Font(
        size=16,
        bold=True,
        color="FFFFFF"
    )
    summary["A1"].fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )
    summary["A1"].alignment = Alignment(
        horizontal="center"
    )

    summary["A3"] = "Server"
    summary["B3"] = guild_name
    summary["A4"] = "Generated At (Chile)"
    summary["B4"] = generated_at_text

    summary_rows = [
        ("Members analyzed", total_members),
        ("Verified", verified_count),
        ("Not verified", not_verified_count),
        ("Has character application", has_application_count),
        ("No character application", no_application_count),
        ("Missing verification + no application", both_missing_count),
        ("Incomplete onboarding (either missing)", incomplete_count),
        ("Has accepted character application", accepted_count),
        ("Approved Player role", approved_player_count),
        ("Members with anomalies", anomaly_member_count),
        ("Total anomaly flags", anomaly_flag_count),
    ]

    start_row = 7

    summary.cell(
        row=start_row,
        column=1,
        value="Metric"
    )
    summary.cell(
        row=start_row,
        column=2,
        value="Count"
    )

    for offset, (label, value) in enumerate(
        summary_rows,
        start=1
    ):
        summary.cell(
            row=start_row + offset,
            column=1,
            value=label
        )
        summary.cell(
            row=start_row + offset,
            column=2,
            value=value
        )

    for cell in summary[start_row]:
        if cell.column <= 2:
            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )
            cell.font = Font(
                color="FFFFFF",
                bold=True
            )

    summary.column_dimensions["A"].width = 42
    summary.column_dimensions["B"].width = 20
    summary.column_dimensions["C"].width = 4
    summary.column_dimensions["D"].width = 4

    chart = BarChart()
    chart.title = "Onboarding Gaps"
    chart.y_axis.title = "Members"
    chart.x_axis.title = "Status"
    chart.height = 7
    chart.width = 11

    # Chart the three main gaps: not verified, no application, both missing.
    chart_data_rows = [
        start_row + 3,
        start_row + 5,
        start_row + 6,
    ]

    # Place compact helper data for the chart in D7:E10.
    summary["D7"] = "Gap"
    summary["E7"] = "Count"
    summary["D8"] = "Not verified"
    summary["E8"] = not_verified_count
    summary["D9"] = "No application"
    summary["E9"] = no_application_count
    summary["D10"] = "Both missing"
    summary["E10"] = both_missing_count

    chart_data = Reference(
        summary,
        min_col=5,
        min_row=7,
        max_row=10
    )
    chart_categories = Reference(
        summary,
        min_col=4,
        min_row=8,
        max_row=10
    )
    chart.add_data(
        chart_data,
        titles_from_data=True
    )
    chart.set_categories(chart_categories)
    summary.add_chart(chart, "D12")

    headers = [
        "Discord Username",
        "Display Name",
        "User ID",
        "Joined Server",
        "Member Role",
        "Verified",
        "Community Access",
        "Character Application",
        "Latest Application Status",
        "Application Count",
        "Latest Application",
        "Latest Application Link",
        "Accepted Application",
        "Approved Player",
        "Race Role(s)",
        "Anomaly Count",
        "Anomaly Severity",
        "Anomaly Type(s)",
        "Needs Action",
        "Missing Verification + No Application",
    ]

    exact_rows = [
        row
        for row in rows
        if row["Missing Verification + No Application"] == "Yes"
    ]

    incomplete_rows = [
        row
        for row in rows
        if (
            row["Verified"] == "No"
            or row["Character Application"] == "No"
        )
    ]

    severity_order = {
        "Critical": 0,
        "High": 1,
        "Review": 2,
        "Clean": 3,
    }

    anomaly_rows = [
        row
        for row in rows
        if row.get("Anomaly Count", 0) > 0
    ]

    anomaly_rows.sort(
        key=lambda item: (
            severity_order.get(
                item.get("Anomaly Severity", "Clean"),
                99
            ),
            -item.get("Anomaly Count", 0),
            item["Display Name"].casefold(),
        )
    )

    write_data_sheet(
        workbook,
        "No Verify + No App",
        exact_rows,
        headers,
        "NoVerifyNoAppTable"
    )

    write_data_sheet(
        workbook,
        "Incomplete Onboarding",
        incomplete_rows,
        headers,
        "IncompleteOnboardingTable"
    )

    anomalies_sheet = write_data_sheet(
        workbook,
        "⚠️ Anomalies",
        anomaly_rows,
        headers,
        "OnboardingAnomaliesTable"
    )

    if "Anomaly Severity" in headers:
        severity_column = headers.index(
            "Anomaly Severity"
        ) + 1

        severity_fills = {
            "Critical": PatternFill(
                "solid",
                fgColor="F4CCCC"
            ),
            "High": PatternFill(
                "solid",
                fgColor="FCE5CD"
            ),
            "Review": PatternFill(
                "solid",
                fgColor="FFF2CC"
            ),
        }

        for row_number in range(
            2,
            anomalies_sheet.max_row + 1
        ):
            cell = anomalies_sheet.cell(
                row=row_number,
                column=severity_column
            )

            fill = severity_fills.get(cell.value)

            if fill is not None:
                cell.fill = fill
                cell.font = Font(bold=True)

    write_data_sheet(
        workbook,
        "All Members",
        rows,
        headers,
        "AllMembersTable"
    )

    workbook.save(output_path)

    return {
        "path": output_path,
        "total": total_members,
        "both_missing": both_missing_count,
        "incomplete": incomplete_count,
        "not_verified": not_verified_count,
        "no_application": no_application_count,
        "anomaly_members": anomaly_member_count,
        "anomaly_flags": anomaly_flag_count,
    }


async def generate_onboarding_report(
    guild: discord.Guild
):
    """Collects live Discord data and writes a timestamped Excel report."""

    application_index = await collect_character_application_index(
        guild
    )

    members = await get_complete_member_list(
        guild
    )

    rows = build_onboarding_rows(
        guild,
        members,
        application_index
    )

    now_chile = datetime.now(
        ZoneInfo(REPORT_TIMEZONE)
    )

    generated_at_text = now_chile.strftime(
        "%Y-%m-%d %H:%M:%S %Z"
    )

    filename = (
        "GoldenAge_Onboarding_Report_"
        f"{now_chile.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    )

    output_path = os.path.join(
        REPORTS_DIR,
        filename
    )

    result = await asyncio.to_thread(
        create_onboarding_excel,
        guild.name,
        rows,
        output_path,
        generated_at_text
    )

    return result


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

def extract_character_race_fallback(content: str):
    """
    Extract Race/Species from old and new Discord application formatting.

    Supported examples include:
      Race: Namekian
      **Race:** Saiyan
      __Species__: Arcosian
      > Race - Bio-Android
      Race = Human (Android)
      Race: Mutant Majin

    This parser only reads a labeled Race/Species field; it does not guess a
    race from the backstory.
    """

    if not content:
        return None

    label = r"(?:race|species|race\s*/\s*species|species\s*/\s*race)"
    separator = r"(?:[:：=]|[-–—])"

    # Prefer a labeled field at the start of a line. Markdown around the label
    # and value is allowed, as are quote/list markers.
    line_pattern = re.compile(
        rf"(?im)^[\s>*#\-•|`~_]*"
        rf"(?:\*{{0,2}}|_{{0,2}})"
        rf"{label}"
        rf"(?:\*{{0,2}}|_{{0,2}})"
        rf"\s*{separator}\s*"
        rf"(?:\*{{0,2}}|_{{0,2}})?"
        rf"([^\r\n]+)"
    )

    match = line_pattern.search(content)

    # Some legacy posts place several fields on one physical line.
    if match is None:
        inline_pattern = re.compile(
            rf"(?i)(?<!\w){label}\s*{separator}\s*"
            rf"(?:\*{{0,2}}|_{{0,2}})?"
            rf"(.+?)"
            rf"(?=\s+(?:birth\s*name|character\s*name|name|alignment|class|"
            rf"combat\s*role|age|sex|gender|family|alias|nickname|"
            rf"fighting\s*style|combat\s*style|mutant(?:\s*status)?|"
            rf"mutation(?:\s*status)?)\s*(?:[:：=]|[-–—])|$)"
        )
        match = inline_pattern.search(content.replace("\n", " "))

    # Very old forms occasionally omit punctuation: "Race Namekian".
    if match is None:
        loose_line_pattern = re.compile(
            rf"(?im)^[\s>*#\-•|`~_]*"
            rf"(?:\*{{0,2}}|_{{0,2}})"
            rf"{label}"
            rf"(?:\*{{0,2}}|_{{0,2}})"
            rf"\s+"
            rf"(?:\*{{0,2}}|_{{0,2}})?"
            rf"([^\r\n]+)"
        )
        match = loose_line_pattern.search(content)

    if match is None:
        return None

    value = match.group(1).strip()
    value = re.sub(r"[\s*_`~]+$", "", value).strip()

    return value or None


def extract_mutant_field_fallback(content: str):
    """Returns the explicit Mutant/Mutation field value when present."""

    if not content:
        return None

    label = (
        r"(?:is\s+(?:this\s+character\s+)?a\s+mutant\??|"
        r"mutant(?:\s*status)?\??|"
        r"mutation(?:\s*status|\s*type)?\??)"
    )

    pattern = re.compile(
        rf"(?im)^[\s>*#\-•|`~_]*"
        rf"(?:\*{{0,2}}|_{{0,2}})"
        rf"{label}"
        rf"(?:\*{{0,2}}|_{{0,2}})"
        rf"\s*(?:[:：=]|[-–—])\s*"
        rf"(?:\*{{0,2}}|_{{0,2}})?"
        rf"([^\r\n]+)"
    )

    match = pattern.search(content)
    if match is None:
        return None

    value = match.group(1).strip()
    value = re.sub(r"[\s*_`~]+$", "", value).strip()
    return value or None


def detect_mutant_status(content: str, application: dict | None = None) -> bool:
    """
    Detects whether an Accepted character is explicitly marked as Mutant.

    Safe signals:
      - Race/Species contains "Mutant" (e.g. "Mutant Saiyan")
      - A labeled Mutant/Mutation field has an affirmative/non-negative value

    Backstory mentions of the word "mutant" are intentionally ignored.
    """

    application = application or {}

    race_value = application.get("race") or ""
    race_normalized = normalize(race_value)

    if re.search(r"\bmutant\b", race_normalized):
        return True

    explicit_value = application.get("mutant")
    if not explicit_value:
        explicit_value = extract_mutant_field_fallback(content)

    if not explicit_value:
        return False

    value = normalize(explicit_value)

    negative_patterns = (
        r"^no$",
        r"^false$",
        r"^0$",
        r"^none$",
        r"^n/?a$",
        r"^not\s+a\s+mutant$",
        r"^non[-\s]?mutant$",
        r"^normal$",
        r"^standard$",
    )

    if any(re.search(pattern, value) for pattern in negative_patterns):
        return False

    positive_patterns = (
        r"^yes\b",
        r"^true\b",
        r"^1\b",
        r"^y\b",
        r"^si\b",
        r"^sí\b",
        r"\bmutant\b",
        r"\bmutation\b",
        r"^active\b",
        r"^approved\b",
    )

    if any(re.search(pattern, value) for pattern in positive_patterns):
        return True

    # A non-empty value in a field explicitly labeled Mutant/Mutation (for
    # example "Mutation Type: Alpha") is treated as a positive designation.
    return True


def mutant_display_name(application: dict, author) -> str:
    """Returns a compact character name for the public Mutant Count."""

    value = clean_value(application.get("character_name") or "")
    value = re.sub(r"[*_`~]", "", value).strip()

    if value:
        value = value.splitlines()[0].strip()

    if not value or len(value) > 40:
        value = getattr(author, "display_name", None) or author.name

    return value[:40]


def parse_character_application(
    content: str
):

    aliases = {

        "character_name": [
            r"birth\s*name",
            r"character\s*name",
            r"character",
            r"name",
        ],

        "race": [
            r"race\s*/\s*species",
            r"species\s*/\s*race",
            r"race",
            r"species",
        ],

        "mutant": [
            r"is\s+(?:this\s+character\s+)?a\s+mutant\??",
            r"mutant\s*status\??",
            r"mutation\s*status\??",
            r"mutation\s*type\??",
            r"mutant\??",
            r"mutation\??",
        ],

        "alias": [
            r"alias",
            r"nickname",
        ],

        "class": [
            r"combat\s*role",
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

    application = parse_labeled_fields(
        content,
        aliases
    )

    if not application.get("race"):
        fallback_race = extract_character_race_fallback(content)

        if fallback_race:
            application["race"] = fallback_race
            print(
                f"🧬 Race parser fallback detected: "
                f"{fallback_race!r}"
            )

    return application


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

        "mutant-application":
            MUTANT_APPLICATION_FORUM_ID,

        "approved-mutants":
            APPROVED_MUTANTS_FORUM_ID,
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

    global INITIAL_ACCESS_SYNC_DONE

    if not INITIAL_ACCESS_SYNC_DONE:
        await rebuild_approved_character_cache()
        await rebuild_approved_mutant_cache()
        await update_race_count_message(reason="startup")
        await sync_all_community_access()
        INITIAL_ACCESS_SYNC_DONE = True
        print()

    print("=" * 65)

    print(
        "🤖 Waiting for applications and role changes..."
    )

    print(
        "📊 Staff report command: !onboardingreport"
    )

    print("=" * 65)
    print()


# ============================================================
# MEMBER / ACCESS AUTOMATION
# ============================================================

@client.event
async def on_member_update(
    before: discord.Member,
    after: discord.Member
):
    """Keeps Community Access aligned with Member + Verified."""

    watched_role_ids = {
        MEMBER_ROLE_ID,
        VERIFIED_ROLE_ID,
        COMMUNITY_ACCESS_ROLE_ID,
    }

    before_ids = {
        role.id
        for role in before.roles
        if role.id in watched_role_ids
    }

    after_ids = {
        role.id
        for role in after.roles
        if role.id in watched_role_ids
    }

    if before_ids == after_ids:
        return

    sync_key = (after.guild.id, after.id)

    if sync_key in ACCESS_SYNC_IN_PROGRESS:
        return

    await sync_community_access_for_member(
        after,
        reason=(
            "Golden Age Manager: Member + Verified access gate"
        )
    )


# ============================================================
# STAFF REPORT COMMAND
# ============================================================

@client.event
async def on_message(message: discord.Message):
    """
    Staff-only manual Excel report.

    Commands:
      !onboardingreport
      !onboarding-report
    """

    if message.author.bot:
        return

    if message.guild is None:
        return

    command = normalize(
        message.content or ""
    )

    if command not in REPORT_COMMANDS:
        return

    # The onboarding report is intentionally owner-only because it
    # contains internal member/onboarding audit data.
    if message.author.id != message.guild.owner_id:
        await message.reply(
            "❌ Only the server owner can generate the onboarding report."
        )
        return

    report_key = (
        message.guild.id,
        message.author.id
    )

    if report_key in REPORT_IN_PROGRESS:
        await message.reply(
            "⏳ An onboarding report is already being generated for you."
        )
        return

    REPORT_IN_PROGRESS.add(report_key)

    progress_message = None

    try:
        progress_message = await message.reply(
            "📊 Generating the onboarding Excel report..."
        )

        result = await generate_onboarding_report(
            message.guild
        )

        summary_text = (
            "## 📊 Onboarding Report Ready\n\n"
            f"👥 **Members analyzed:** `{result['total']}`\n"
            f"❌ **Not verified:** `{result['not_verified']}`\n"
            f"📝 **No character application:** `{result['no_application']}`\n"
            f"🚨 **Missing both verification + application:** "
            f"`{result['both_missing']}`\n"
            f"⚠️ **Incomplete onboarding:** `{result['incomplete']}`\n"
            f"🧭 **Members with anomalies:** `{result['anomaly_members']}`\n"
            f"🚩 **Total anomaly flags:** `{result['anomaly_flags']}`\n\n"
            "The Excel file was also saved in the bot's `reports/` folder."
        )

        await message.channel.send(
            summary_text,
            file=discord.File(
                result["path"],
                filename=os.path.basename(result["path"])
            )
        )

        if progress_message is not None:
            try:
                await progress_message.delete()
            except discord.HTTPException:
                pass

        print(
            f"📊 Onboarding report generated by {message.author}: "
            f"{result['path']}"
        )

    except Exception as error:
        print(
            f"❌ Onboarding report generation failed: {error!r}"
        )

        error_text = (
            "❌ **Golden Age Manager could not generate the onboarding "
            "report.**\n\n"
            f"`{type(error).__name__}: {error}`"
        )

        if progress_message is not None:
            try:
                await progress_message.edit(
                    content=error_text
                )
            except discord.HTTPException:
                await message.channel.send(error_text)
        else:
            await message.channel.send(error_text)

    finally:
        REPORT_IN_PROGRESS.discard(report_key)


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


    if (
        after.parent_id
        == MUTANT_APPLICATION_FORUM_ID
    ):

        await handle_mutant_application(
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

    accepted_normalized = normalize(
        CHARACTER_ACCEPTED_TAG
    )

    accepted_removed = (
        accepted_normalized in before_tags
        and accepted_normalized not in after_tags
    )

    if accepted_removed:
        await rebuild_approved_character_cache()
        await update_race_count_message(
            reason="Accepted tag removed"
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

        await rebuild_approved_character_cache()
        await update_race_count_message(
            reason="character rejected"
        )

        return


    # --------------------------------------------------------
    # ACCEPTED
    # --------------------------------------------------------

    if accepted_normalized not in newly_added:
        return


    # --------------------------------------------------------
    # DUPLICATE PROTECTION
    # --------------------------------------------------------

    if application_was_processed(
        "character",
        after.id
    ):

        destination_exists = await processed_destination_exists(
            "character",
            after.id,
            APPROVED_CHARACTERS_FORUM_ID
        )

        if destination_exists:
            # The approved profile still exists, so do not duplicate it.
            # The Accepted tag may simply have been restored after a test
            # or staff correction.
            await rebuild_approved_character_cache()
            await update_race_count_message(
                reason="Accepted tag restored on processed application"
            )

            print(
                "⚠️ This character application was already processed. "
                "Profile duplicate skipped; Race Count refreshed."
            )

            return

        # The source application is marked as processed, but its approved
        # profile no longer exists (for example, Staff manually deleted it
        # from approved-characters during a test). Clear the stale marker
        # and continue through the normal approval flow so the profile is
        # recreated automatically.
        unmark_application_processed(
            "character",
            after.id
        )

        print(
            "♻️ Processed character profile is missing from "
            "approved-characters. Rebuilding it now..."
        )


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


    member = await get_guild_member(
        after.guild,
        author.id
    )

    if member is None:

        print(
            f"❌ Could not resolve applicant as a guild member: "
            f"{author} ({author.id})"
        )

        try:
            await after.send(
                "⚠️ **Golden Age Manager could not resolve the "
                "applicant as a server member.**\n\n"
                "The application was not processed."
            )
        except discord.HTTPException:
            pass

        return


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
                "`Race: Human`\n"
                "`Race: Human (Android)`\n"
                "`Race: Namekian`\n"
                "`Race: Arcosian`\n"
                "`Race: Bio-Android`\n"
                "`Race: Majin`\n"
                "Mutants can use formats such as "
                "`Race: Mutant Saiyan` or `Mutant: Yes`.\n\n"
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


    _, _, access_role_error = resolve_character_access_roles(
        after.guild,
        race_tag_name
    )

    if access_role_error:

        print(
            f"❌ Character access role validation failed: "
            f"{access_role_error}"
        )

        try:
            await after.send(
                "⚠️ **Character approval is ready, but Golden Age "
                "Manager cannot assign the required Discord roles.**\n\n"
                f"`{access_role_error}`\n\n"
                "The application was **not published or archived**. "
                "Please fix the bot role permissions/hierarchy and "
                "apply the Accepted tag again."
            )
        except discord.HTTPException:
            pass

        return


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


        APPROVED_CHARACTER_CACHE[
            member.id
        ] = race_tag_name


        access_ok, access_result = (
            await grant_character_access_roles(
                member,
                race_tag_name
            )
        )


        mark_application_processed(
            "character",
            after.id,
            approved_thread.id
        )


        # Re-scan Accepted applications so the public counter is based on
        # accepted CHARACTER THREADS rather than unique Discord users.
        await rebuild_approved_character_cache()

        await update_race_count_message(
            reason=f"approved {display_name}"
        )


        print(
            f"✅ Approved character created: "
            f"{approved_thread.name}"
        )


        if access_ok:
            access_line = (
                f"🎮 **Server Access:** Granted (`{access_result}`)"
            )
        else:
            access_line = (
                "⚠️ **Server Access:** Could not be assigned "
                f"automatically (`{access_result}`). Staff has been "
                "notified."
            )


        await after.send(
            "## ✅ Character Approved\n\n"
            "Your character application has "
            "been accepted.\n\n"
            f"🧬 **Race:** "
            f"`{race_tag_name}`\n\n"
            f"{access_line}\n\n"
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
# MUTANT AUTOMATION
# ============================================================

async def handle_mutant_application(
    before: discord.Thread,
    after: discord.Thread,
):
    """Processes the dedicated Mutant Application forum."""

    before_tags = {normalize(tag.name) for tag in before.applied_tags}
    after_tags = {normalize(tag.name) for tag in after.applied_tags}
    newly_added = after_tags - before_tags

    accepted_name = normalize(MUTANT_ACCEPTED_TAG)
    rejected_name = normalize(MUTANT_REJECTED_TAG)

    accepted_removed = (
        accepted_name in before_tags
        and accepted_name not in after_tags
    )

    if accepted_removed:
        await rebuild_approved_mutant_cache()
        await update_race_count_message(
            reason="Mutant Accepted tag removed"
        )

    if not newly_added:
        return

    if rejected_name in newly_added:
        await reject_application(
            after,
            "Mutant Application"
        )
        await rebuild_approved_mutant_cache()
        await update_race_count_message(
            reason="mutant application rejected"
        )
        return

    if accepted_name not in newly_added:
        return

    # --------------------------------------------------------
    # MUTANT CAPACITY GUARD
    # --------------------------------------------------------
    # Rebuild after the Accepted tag is added so the candidate is included
    # in the source-of-truth count. If this approval would make the race
    # exceed 3/3, Golden Age Manager removes Accepted and leaves the
    # application open for Staff to resolve.
    await rebuild_approved_mutant_cache()

    candidate_record = APPROVED_MUTANT_RECORDS.get(after.id)

    if candidate_record is not None:
        candidate_race = candidate_record.get("race")
        same_race_records = [
            record
            for record in APPROVED_MUTANT_RECORDS.values()
            if record.get("race") == candidate_race
        ]
        candidate_count = len(same_race_records)

        if candidate_count > MUTANT_LIMIT_PER_RACE:
            display_race = MUTANT_COUNT_RACES.get(
                candidate_race,
                candidate_race or "Unknown",
            )

            try:
                remaining_tags = [
                    tag
                    for tag in after.applied_tags
                    if normalize(tag.name) != accepted_name
                ]

                await after.edit(
                    applied_tags=remaining_tags,
                    reason=(
                        "Golden Age Manager: Mutant race capacity reached"
                    ),
                )

                await after.send(
                    "## 🚫 Mutant Capacity Reached\n\n"
                    f"**{display_race} Mutants are currently full:** "
                    f"`{MUTANT_LIMIT_PER_RACE}/{MUTANT_LIMIT_PER_RACE}`.\n\n"
                    "This application cannot be approved until a slot "
                    "becomes available. Golden Age Manager removed the "
                    "`Accepted` tag automatically.\n\n"
                    "The application has **not** been published or archived."
                )

            except discord.Forbidden:
                print(
                    "❌ Could not remove Accepted after Mutant capacity "
                    "was exceeded. Check Manage Threads/forum permissions."
                )

            except discord.HTTPException as error:
                print(
                    f"❌ Discord API error while enforcing Mutant cap: {error}"
                )

            await rebuild_approved_mutant_cache()
            await update_race_count_message(
                reason="Mutant approval blocked by race capacity"
            )

            print(
                f"🚫 Mutant approval blocked: {display_race} is "
                f"{MUTANT_LIMIT_PER_RACE}/{MUTANT_LIMIT_PER_RACE} FULL."
            )
            return

    # --------------------------------------------------------
    # DUPLICATE / SELF-HEAL PROTECTION
    # --------------------------------------------------------

    if application_was_processed("mutant", after.id):
        destination_exists = await processed_destination_exists(
            "mutant",
            after.id,
            APPROVED_MUTANTS_FORUM_ID,
        )

        if destination_exists:
            await rebuild_approved_mutant_cache()
            await update_race_count_message(
                reason="Mutant Accepted tag restored on processed application"
            )

            print(
                "⚠️ This mutant application was already processed. "
                "Profile duplicate skipped; Mutant Count refreshed."
            )
            return

        unmark_application_processed("mutant", after.id)
        print(
            "♻️ Processed mutant profile is missing from approved-mutants. "
            "Rebuilding it now..."
        )

    print()
    print("=" * 65)
    print("🧬 MUTANT APPROVAL DETECTED")
    print(f"📋 {after.name}")
    print("=" * 65)

    approved_forum = await get_channel(APPROVED_MUTANTS_FORUM_ID)
    if not isinstance(approved_forum, discord.ForumChannel):
        print("❌ approved-mutants not found.")
        return

    starter_message = await get_starter_message(after)
    if starter_message is None:
        print("❌ Mutant starter message not found.")
        return

    mutant_application = parse_mutant_application(
        starter_message.content or ""
    )

    character_record, error = await resolve_character_application_for_mutant(
        starter_message.author.id,
        mutant_application,
    )

    if error or character_record is None:
        print(f"❌ Mutant approval blocked: {error}")

        try:
            await after.send(
                "⚠️ **Golden Age Manager could not validate this Mutant "
                "Application.**\n\n"
                f"{error}\n\n"
                "Please make sure the applicant has an **Accepted Character "
                "Application** and that the Mutant form contains its exact "
                "Discord link in `Character Application:`.\n\n"
                "The application was **not published or archived**."
            )
        except discord.HTTPException:
            pass

        return

    character_race = character_record["race"]
    destination_tag_name = MUTANT_APPROVED_TAG_BY_CHARACTER_RACE.get(
        character_race
    )

    if destination_tag_name is None:
        print(
            f"❌ Mutant race {character_race!r} has no approved-mutants tag mapping."
        )
        return

    destination_tag = find_forum_tag(
        approved_forum,
        destination_tag_name,
    )

    if destination_tag is None:
        print(
            f"❌ Destination mutant tag '{destination_tag_name}' "
            "does not exist in approved-mutants."
        )
        return

    character_name = (
        clean_value(mutant_application.get("character_name") or "")
        or clean_value(character_record.get("name") or "")
        or after.name
    )
    character_name = re.sub(r"[*_`~]", "", character_name)
    character_name = character_name.splitlines()[0].strip()

    normalized_after_tags = {normalize(tag.name) for tag in after.applied_tags}

    if normalize(MUTANT_MIGRATION_TAG) in normalized_after_tags:
        application_type = "Migration"
    elif normalize(MUTANT_NEW_APPLICATION_TAG) in normalized_after_tags:
        application_type = "New Application"
    else:
        application_type = clean_value(
            mutant_application.get("application_type") or ""
        ) or "Unspecified"

    embed = discord.Embed(
        title="✅ Approved Mutant"
    )

    embed.set_author(
        name=starter_message.author.display_name,
        icon_url=starter_message.author.display_avatar.url,
    )

    embed.description = (
        f"## {character_name}\n"
        f"**{destination_tag_name} Mutant**"
    )

    embed.add_field(
        name="🧬 Race",
        value=destination_tag_name,
        inline=True,
    )

    embed.add_field(
        name="📦 Application Type",
        value=application_type[:1024],
        inline=True,
    )

    fields = [
        ("🧬 Mutation Name / Type", "mutation_name"),
        ("🌌 Origin of the Mutation", "origin"),
        ("👁️ Physical Changes", "physical_changes"),
        ("⚡ Abilities / Advantages", "abilities"),
        ("⚠️ Weaknesses / Limitations", "weaknesses"),
        ("📖 Mutation Backstory", "mutation_backstory"),
        ("🎭 Effect on Character", "character_effect"),
    ]

    for label, key in fields:
        value = mutant_application.get(key)
        if value:
            add_long_field(embed, label, value)

    if application_type.casefold() == "migration":
        migration_lines = []
        if mutant_application.get("previously_approved"):
            migration_lines.append(
                f"**Previously Approved:** {mutant_application['previously_approved']}"
            )
        if mutant_application.get("previous_character_name"):
            migration_lines.append(
                f"**Previous Character Name:** {mutant_application['previous_character_name']}"
            )
        if mutant_application.get("previous_mutant_race"):
            migration_lines.append(
                f"**Previous Mutant Race:** {mutant_application['previous_mutant_race']}"
            )

        if migration_lines:
            add_long_field(
                embed,
                "📦 Migration Information",
                "\n".join(migration_lines),
            )

    proof = mutant_application.get("proof")
    if proof:
        add_long_field(embed, "🧾 Proof / Evidence", proof)

    additional_info = mutant_application.get("additional_info")
    if additional_info:
        add_long_field(embed, "➕ Additional Information", additional_info)

    embed.add_field(
        name="🔗 Character Application",
        value=(
            f"[View approved character application]"
            f"({character_record['starter'].jump_url})"
        ),
        inline=False,
    )

    embed.add_field(
        name="🔗 Mutant Application",
        value=(
            f"[View original mutant application]"
            f"({starter_message.jump_url})"
        ),
        inline=False,
    )

    embed.set_footer(
        text="Dragon Ball: Golden Age • Mutant Registry"
    )

    files, thumbnail_filename = await prepare_attachments(starter_message)
    if thumbnail_filename:
        embed.set_thumbnail(
            url=f"attachment://{thumbnail_filename}"
        )

    try:
        result = await approved_forum.create_thread(
            name=character_name[:100],
            embed=embed,
            applied_tags=[destination_tag],
            files=files,
        )

        approved_thread = result.thread

        mark_application_processed(
            "mutant",
            after.id,
            approved_thread.id,
        )

        print(
            f"✅ Approved mutant profile created: {approved_thread.name}"
        )

        await after.send(
            "## ✅ Mutant Application Approved\n\n"
            "Your Mutant Application has been accepted.\n\n"
            f"🧬 **Character:** `{character_name}`\n"
            f"🌌 **Mutant Race:** `{destination_tag_name}`\n"
            f"📦 **Type:** `{application_type}`\n\n"
            "Your Mutant profile has been published in "
            f"{approved_forum.mention}."
        )

        await after.edit(
            archived=True,
            locked=True,
            reason="Golden Age Manager: Mutant approved",
        )

        await rebuild_approved_mutant_cache()
        await update_race_count_message(
            reason="mutant application approved"
        )

        print("🔒 Original mutant application archived.")
        print("✅ Mutant automation completed successfully.")
        print()

    except discord.Forbidden:
        print(
            "❌ MUTANT PERMISSION ERROR — Golden Age Manager does not "
            "have enough permissions in the Mutant forums."
        )

    except discord.HTTPException as error:
        print(f"❌ MUTANT DISCORD API ERROR: {error}")

    except Exception as error:
        print(f"❌ MUTANT UNEXPECTED ERROR: {repr(error)}")


# ============================================================
# START GOLDEN AGE MANAGER
# ============================================================

client.run(TOKEN)