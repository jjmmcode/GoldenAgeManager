# Golden Age Manager v8.1
# Adds DragonMineZ 2.2 techniques to the Mentor Ability catalogue.
# Preserves Grey/Demon races, departed-member cleanup and active roster synchronization.
import os
import re
import json
import asyncio
import tempfile
import unicodedata
import zipfile
from datetime import datetime
from zoneinfo import ZoneInfo

import discord
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from docx import Document
from docx.shared import Inches, Pt, RGBColor


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

# Public mentor directory message. Uses the same #information channel as
# Race Count so players can quickly see who currently mentors each type.
MENTOR_LIST_CHANNEL_ID = 1543938120301551626

# Source-of-truth records: one record per currently approved Mentor Application.
# source_thread_id -> {"user_id", "character", "skill", "type", "thread_name"}
APPROVED_MENTOR_RECORDS = {}
MENTOR_LIST_LOCK = asyncio.Lock()

# Public catalogue of mentor slots that are still open. Unlike the Mentor
# List, this is NOT built from approved applications. It starts from a fixed
# whitelist of abilities that Staff allows people to apply for, then removes
# any ability that already has an approved mentor.
AVAILABLE_ABILITIES_CHANNEL_ID = MENTOR_LIST_CHANNEL_ID
AVAILABLE_ABILITIES_LOCK = asyncio.Lock()

# Source-of-truth catalogue for OPEN mentor applications.
# Custom Forms / Passive Skills stay empty until Staff provides their lists;
# empty categories are omitted from the public message instead of inventing
# availability from already-approved Mentor Applications.
MENTOR_ABILITY_CATALOG = {
    "Technique": [
        # Existing Golden Age mentor catalogue
        "Big Bang",
        "Burning Attack",
        # Death Beam + Emperor Death Beam share ONE mentor slot.
        # A mentor must know both techniques to occupy this slot.
        "Death Beam + Emperor Death Beam",
        "Fake Moon",
        "Final Explosion",
        "Final Flash",
        "Galick Gun",
        "Kamehameha",
        "Ki Barrage",
        # Kienzan + Kienzan Doble share ONE mentor slot.
        # A mentor must know both techniques to occupy this slot.
        "Kienzan + Kienzan Doble",
        "Makkanko",
        "Masenko",
        "Sokidan",
        "Soul Punisher",
        "Spirit Bomb",
        "Supernova",
        "Supernova Cooler",

        # DragonMineZ 2.2 — new Ki techniques
        "Heat Dome Attack",
        "Hellzone Grenade",
        "Dodon Ray",
        "Tri-Beam",
        "Neo Tri-Beam",
        "Assault Rain",
        "Blaster Meteor",
        "Mouth Blast",
        "Kill Driver",

        # DragonMineZ 2.2 — new Strike techniques
        "Spirit Breaking Cannon",
        "Shining Sword Attack",
        "Oozaru Slam",
        "Gum Punch",
        "Blue Hurricane",
        "Grab Throw",
        "Meteor Combo",
        "Meteor Burst",
        "Mystic Ball Attack",
        "Dimensional Fist",
        "Dimensional Slash",

        # DragonMineZ 2.2 — Evasion techniques
        # Solar Flare existed previously but moved into the new Evasion
        # category in v2.2; it was not present in Golden Age's old whitelist.
        "Afterimage",
        "Solar Flare",
        "Rage Scream",
        "Sleep Recovery",
        "Dimensional Warp",
    ],
    "Custom-Form": [],
    "Passive-Skill": [],
}

# Common wording variants accepted when matching an approved application to
# the catalogue. The values must be exact names from MENTOR_ABILITY_CATALOG.
MENTOR_ABILITY_ALIASES = {
    "big bang": "Big Bang",
    "big bang attack": "Big Bang",
    "burning attack": "Burning Attack",
    # Death Beam is intentionally NOT mapped individually here.
    # It is a paired mentor slot with Emperor Death Beam and is validated below.
    "fake moon": "Fake Moon",
    "power ball": "Fake Moon",
    "final explosion": "Final Explosion",
    "final flash": "Final Flash",
    "galick gun": "Galick Gun",
    "galick gun attack": "Galick Gun",
    "kamehameha": "Kamehameha",
    "ki barrage": "Ki Barrage",
    # Kienzan is intentionally NOT mapped individually here.
    # It is a paired mentor slot with Kienzan Doble and is validated below.
    "makkanko": "Makkanko",
    "makankosappo": "Makkanko",
    "special beam cannon": "Makkanko",
    "masenko": "Masenko",
    "sokidan": "Sokidan",
    "spirit ball": "Sokidan",
    "soul punisher": "Soul Punisher",
    "spirit bomb": "Spirit Bomb",
    "spiritbomb": "Spirit Bomb",
    "genki dama": "Spirit Bomb",
    "supernova": "Supernova",
    "supernova cooler": "Supernova Cooler",
    "cooler supernova": "Supernova Cooler",

    # DragonMineZ 2.2 — Ki techniques
    "heat dome attack": "Heat Dome Attack",
    "heat dome": "Heat Dome Attack",
    "hellzone grenade": "Hellzone Grenade",
    "hell zone grenade": "Hellzone Grenade",
    "dodon ray": "Dodon Ray",
    "dodonpa": "Dodon Ray",
    "dodon pa": "Dodon Ray",
    "tri-beam": "Tri-Beam",
    "tri beam": "Tri-Beam",
    "tribeam": "Tri-Beam",
    "kikoho": "Tri-Beam",
    "neo tri-beam": "Neo Tri-Beam",
    "neo tri beam": "Neo Tri-Beam",
    "neo tribeam": "Neo Tri-Beam",
    "neo kikoho": "Neo Tri-Beam",
    "assault rain": "Assault Rain",
    "blaster meteor": "Blaster Meteor",
    "mouth blast": "Mouth Blast",
    "kill driver": "Kill Driver",

    # DragonMineZ 2.2 — Strike techniques
    "spirit breaking cannon": "Spirit Breaking Cannon",
    "spirit-breaking cannon": "Spirit Breaking Cannon",
    "shining sword attack": "Shining Sword Attack",
    "shining sword": "Shining Sword Attack",
    "oozaru slam": "Oozaru Slam",
    "great ape slam": "Oozaru Slam",
    "gum punch": "Gum Punch",
    "blue hurricane": "Blue Hurricane",
    "grab throw": "Grab Throw",
    "grab and throw": "Grab Throw",
    "meteor combo": "Meteor Combo",
    "meteor burst": "Meteor Burst",
    "mystic ball attack": "Mystic Ball Attack",
    "mystic ball": "Mystic Ball Attack",
    "dimensional fist": "Dimensional Fist",
    "dimension fist": "Dimensional Fist",
    "dimensional slash": "Dimensional Slash",
    "dimension slash": "Dimensional Slash",

    # DragonMineZ 2.2 — Evasion techniques
    "afterimage": "Afterimage",
    "after image": "Afterimage",
    "zanzoken": "Afterimage",
    "solar flare": "Solar Flare",
    "taiyoken": "Solar Flare",
    "taiyo ken": "Solar Flare",
    "rage scream": "Rage Scream",
    "sleep recovery": "Sleep Recovery",
    "regenerative sleep": "Sleep Recovery",
    "dimensional warp": "Dimensional Warp",
    "dimension warp": "Dimensional Warp",
}


# Kienzan mentorship is a paired requirement: one mentor owns both techniques.
KIENZAN_BUNDLE_CATALOG_NAME = "Kienzan + Kienzan Doble"
KIENZAN_SINGLE_NAMES = (
    "kienzan",
    "destructo disc",
    "destructo disk",
)
KIENZAN_DOUBLE_NAMES = (
    "kienzan doble",
    "kienzan double",
    "double kienzan",
    "double destructo disc",
    "double destructo disk",
    "destructo disc double",
    "destructo disk double",
)

# Death Beam mentorship is also paired: one mentor owns both techniques.
DEATH_BEAM_BUNDLE_CATALOG_NAME = "Death Beam + Emperor Death Beam"
DEATH_BEAM_SINGLE_NAMES = (
    "death beam",
)
EMPEROR_DEATH_BEAM_NAMES = (
    "emperor death beam",
    "emperor's death beam",
    "emperors death beam",
)


# ------------------------------------------------------------
# CHARACTER SYSTEM
# ------------------------------------------------------------

CHARACTER_APPLICATION_FORUM_ID = 1543968641853300827
APPROVED_CHARACTERS_FORUM_ID = 1543909774062518333

CHARACTER_ACCEPTED_TAG = "Accepted"
CHARACTER_REJECTED_TAG = "Rejected"

# Forum where players publish their complete character stories/lore.
CHARACTER_STORIES_CHANNEL_ID = 1543909873278787674


# ------------------------------------------------------------
# FORM APPLICATION SYSTEM
# ------------------------------------------------------------

FORM_APPLICATION_FORUM_ID = 1546894012072329256
APPROVED_FORMS_FORUM_ID = 1547016509619179540

FORM_ACCEPTED_TAG = "Accepted"
FORM_REJECTED_TAG = "Rejected"
FORM_PENDING_TAG = "Pending"
FORM_VOIDED_TAG = "Voided"


# ------------------------------------------------------------
# LEGACY MUTANT SYSTEM (RETIRED IN v7.0)
# ------------------------------------------------------------
# The old direct Mutant application / approval workflow is intentionally
# disabled. Existing constants and helper functions remain below only for
# backwards compatibility with historical data and can be removed in a later
# cleanup once Staff confirms the archived forums are no longer needed.
# v7.0 does NOT scan these forums, publish approved-mutant profiles, enforce
# Mutant race caps, or include a Mutant Count in #information.

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
    "Tuffle": 1544872038534226021,

    # v8.0 — New playable races.
    # Grey Race and Demon Race currently share the same Discord role ID as
    # provided by Staff. Their Character Applications remain the source of
    # truth that distinguishes which Race Count bucket they occupy.
    "Grey Race": 1554845472072536144,
    "Demon Race": 1554845472072536144,
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

# Counts are based on Accepted Character Applications filtered against the
# current Discord roster and matching race roles. Departed players or stale
# race roles no longer consume public race slots.
RACE_COUNT_LIMITS = {
    "Saiyan": ("Saiyan", 11),
    "Human": ("Human", 10),
    "Android": ("Human (Android)", 2),
    "Namekian": ("Namekian", 10),
    "Frost-Demon": ("Arcosian", 10),
    "Bio-Android": ("Bio-Android", 10),
    "Majin": ("Majin", 10),
    "Tuffle": ("Tuffle", 5),
    "Grey Race": ("Grey Race", 3),
    "Demon Race": ("Demon Race", 3),
}

# Legacy-only: retained for historical parsing helpers. v7.0 no longer
# publishes or enforces Mutant counts/caps.
MUTANT_COUNT_RACES = {
    "Saiyan": "Saiyan",
    "Human": "Human",
    "Namekian": "Namekian",
    "Frost-Demon": "Arcosian",
    "Bio-Android": "Bio-Android",
    "Majin": "Majin",
}

# Legacy-only historical value; no Mutant cap is enforced in v7.0.
MUTANT_LIMIT_PER_RACE = 3

RACE_COUNT_LOCK = asyncio.Lock()

# ------------------------------------------------------------
# ONBOARDING REPORT SYSTEM
# ------------------------------------------------------------

REPORT_COMMANDS = {
    "!onboardingreport",
    "!onboarding-report",
}

STORY_REPORT_COMMANDS = {
    "!storyreport",
    "!story-report",
}

STORY_REPORT_ALL_COMMANDS = {
    "!storyreportall",
    "!story-report-all",
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
    "human androids": "Android",
    "human/android": "Android",
    "human/androids": "Android",
    "human-android": "Android",
    "human-androids": "Android",
    "human (android)": "Android",
    "human (androids)": "Android",
    "android human": "Android",
    "android/human": "Android",
    "android-human": "Android",
    "cyborg": "Android",
    "human cyborg": "Android",
    "organic android": "Android",
    "human organic android": "Android",
    "human cyborg organic android": "Android",
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

    # Tuffles / Tsufuru
    "tuffle": "Tuffle",
    "tuffles": "Tuffle",
    "tuffle race": "Tuffle",
    "tsufuru": "Tuffle",
    "tsufurujin": "Tuffle",
    "tsufru": "Tuffle",
    "tsufruian": "Tuffle",
    "tsufurian": "Tuffle",

    # Grey Race / Jiren's race
    "grey race": "Grey Race",
    "gray race": "Grey Race",
    "grey": "Grey Race",
    "gray": "Grey Race",
    "jiren race": "Grey Race",
    "jiren's race": "Grey Race",
    "jirens race": "Grey Race",
    "jiren species": "Grey Race",
    "grey alien": "Grey Race",
    "gray alien": "Grey Race",

    # Demon Race
    "demon race": "Demon Race",
    "demon": "Demon Race",
    "demons": "Demon Race",
    "demon clan": "Demon Race",
    "demon realm race": "Demon Race",
    "demon realm": "Demon Race",
    "demonic race": "Demon Race",
    "demonic": "Demon Race",
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
        "form": {},
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
        data.setdefault("form", {})
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


def normalize_forum_tag(text: str) -> str:
    """
    Normalize Discord forum tag names while ignoring decorative emoji
    prefixes/suffixes such as "✅ Accepted" or "📦 Migration".
    """

    value = normalize(text)

    if not value:
        return ""

    value = re.sub(r"^[^a-z0-9]+", "", value)
    value = re.sub(r"[^a-z0-9]+$", "", value)

    return value.strip()


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

    wanted = normalize_forum_tag(tag_name)

    for tag in forum.available_tags:

        if normalize_forum_tag(tag.name) == wanted:
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


# ============================================================
# DEPARTED MEMBER PUBLICATION CLEANUP
# ============================================================

# Source forums whose user-owned threads should be removed when the owner
# leaves/is kicked from the server. Bot-generated approved mirrors are deleted
# through PROCESSED_APPLICATIONS before the source thread is removed.
MANAGED_DEPARTURE_FORUMS = {
    "mentor": MENTOR_APPLICATION_FORUM_ID,
    "character": CHARACTER_APPLICATION_FORUM_ID,
    "form": FORM_APPLICATION_FORUM_ID,
    "story": CHARACTER_STORIES_CHANNEL_ID,
}


async def _delete_managed_thread(
    thread: discord.Thread,
    *,
    label: str,
) -> bool:
    """Delete one managed Discord thread with safe diagnostics."""

    try:
        await thread.delete(
            reason=f"Golden Age Manager: departed member cleanup ({label})"
        )
        print(
            f"🗑️ Deleted {label}: "
            f"'{thread.name}' ({thread.id})."
        )
        return True

    except discord.NotFound:
        return False

    except discord.Forbidden:
        print(
            f"❌ Could not delete {label} '{thread.name}'. "
            "Check Manage Threads / Manage Channels permissions."
        )
        return False

    except discord.HTTPException as error:
        print(
            f"❌ Discord API error deleting {label} "
            f"'{thread.name}': {error}"
        )
        return False


async def _delete_processed_destination_for_source(
    application_type: str,
    source_thread_id: int,
) -> tuple[bool, bool]:
    """
    Delete the bot-generated approved mirror associated with one source
    application and remove its processed_applications.json record.

    Returns:
        (destination_deleted, state_changed)
    """

    if application_type not in {"mentor", "character", "form"}:
        return False, False

    applications = PROCESSED_APPLICATIONS.setdefault(
        application_type,
        {}
    )

    record = applications.get(
        str(source_thread_id)
    )

    if not isinstance(record, dict):
        return False, False

    destination_deleted = False
    destination_thread_id = record.get(
        "destination_thread_id"
    )

    if destination_thread_id:
        destination = client.get_channel(
            int(destination_thread_id)
        )

        if destination is None:
            try:
                destination = await client.fetch_channel(
                    int(destination_thread_id)
                )
            except (
                discord.NotFound,
                discord.Forbidden,
                discord.HTTPException,
            ):
                destination = None

        if isinstance(destination, discord.Thread):
            destination_deleted = await _delete_managed_thread(
                destination,
                label=f"approved-{application_type} mirror",
            )

    # The source is being removed regardless, so its processed marker must not
    # survive and point to a deleted/nonexistent application.
    applications.pop(
        str(source_thread_id),
        None
    )

    return destination_deleted, True


async def cleanup_departed_member_publications(
    guild: discord.Guild,
    *,
    user_id: int | None = None,
    reason: str = "departed member cleanup",
):
    """
    Delete managed forum publications owned by members who are no longer in
    the server.

    If user_id is provided, only that departed member is scanned (used by
    on_member_remove). If omitted, every managed source forum is audited
    against the current guild roster (used at startup).

    Deleted content:
      - #mentor-application threads
      - their bot-generated #approved-mentors mirrors
      - #character-application threads
      - their bot-generated #approved-characters mirrors
      - #form-application threads
      - their bot-generated #approved-forms mirrors
      - Character Stories forum threads, when the configured stories channel
        is a Discord ForumChannel

    Ordinary chat messages and unrelated channels are never touched.
    """

    stats = {
        "source_threads": 0,
        "approved_mirrors": 0,
        "members": set(),
        "state_changed": False,
    }

    for application_type, channel_id in MANAGED_DEPARTURE_FORUMS.items():
        channel = await get_channel(
            channel_id
        )

        # Character Stories may be configured as a regular text channel.
        # v7.9 intentionally deletes only managed FORUM publications, not
        # arbitrary chat history.
        if not isinstance(channel, discord.ForumChannel):
            if application_type != "story":
                print(
                    f"⚠️ Departure cleanup skipped {application_type}: "
                    "configured channel is not a forum."
                )
            continue

        threads = await collect_forum_threads(
            channel
        )

        for thread in threads:
            starter = await get_starter_message(
                thread
            )

            if starter is None or starter.author.bot:
                continue

            owner_id = starter.author.id

            if user_id is not None:
                if owner_id != user_id:
                    continue
                # on_member_remove already establishes that this user departed.
                owner_is_departed = True
            else:
                member = await get_guild_member(
                    guild,
                    owner_id
                )
                owner_is_departed = (
                    member is None
                )

            if not owner_is_departed:
                continue

            stats["members"].add(
                owner_id
            )

            if application_type in {
                "mentor",
                "character",
                "form",
            }:
                (
                    mirror_deleted,
                    state_changed,
                ) = await _delete_processed_destination_for_source(
                    application_type,
                    thread.id,
                )

                if mirror_deleted:
                    stats["approved_mirrors"] += 1

                if state_changed:
                    stats["state_changed"] = True

            deleted = await _delete_managed_thread(
                thread,
                label=f"{application_type} source publication",
            )

            if deleted:
                stats["source_threads"] += 1

    if stats["state_changed"]:
        save_processed_applications()

    if (
        stats["source_threads"]
        or stats["approved_mirrors"]
    ):
        print(
            f"🧹 Departure cleanup ({reason}): "
            f"{stats['source_threads']} source publication(s) + "
            f"{stats['approved_mirrors']} approved mirror(s) deleted "
            f"for {len(stats['members'])} departed member(s)."
        )
    else:
        print(
            f"🧹 Departure cleanup ({reason}): "
            "no stale managed publications found."
        )

    return {
        "source_threads": stats["source_threads"],
        "approved_mirrors": stats["approved_mirrors"],
        "members": len(stats["members"]),
    }


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

    accepted_name = normalize_forum_tag(MUTANT_ACCEPTED_TAG)
    rejected_name = normalize_forum_tag(MUTANT_REJECTED_TAG)
    accepted_seen = 0
    skipped = 0

    for thread in await collect_forum_threads(forum):
        tag_names = {normalize_forum_tag(tag.name) for tag in thread.applied_tags}

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


async def reconcile_accepted_mutant_applications():
    """
    Self-heals already-Accepted Mutant Applications at startup.

    This is especially useful after changing tag formats (for example, from
    "Accepted" to "✅ Accepted"). Accepted applications that are missing
    their approved-mutants profile are reprocessed automatically.
    """

    forum = await get_channel(MUTANT_APPLICATION_FORUM_ID)

    if not isinstance(forum, discord.ForumChannel):
        return

    accepted_name = normalize_forum_tag(MUTANT_ACCEPTED_TAG)
    rejected_name = normalize_forum_tag(MUTANT_REJECTED_TAG)
    repaired = 0

    class _BeforeState:
        applied_tags = []

    for thread in await collect_forum_threads(forum):
        tag_names = {normalize_forum_tag(tag.name) for tag in thread.applied_tags}

        if accepted_name not in tag_names or rejected_name in tag_names:
            continue

        if application_was_processed("mutant", thread.id):
            destination_exists = await processed_destination_exists(
                "mutant",
                thread.id,
                APPROVED_MUTANTS_FORUM_ID,
            )

            if destination_exists:
                continue

        print(
            f"♻️ Reprocessing Accepted Mutant Application at startup: "
            f"{thread.name}"
        )

        await handle_mutant_application(_BeforeState(), thread)
        repaired += 1

    if repaired:
        print(
            f"🧬 Mutant startup reconciliation processed "
            f"{repaired} Accepted application(s)."
        )


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


def get_race_counts(guild: discord.Guild):
    """
    Returns the CURRENT active race population.

    A character only consumes a slot when:
      1) its Character Application is currently Accepted,
      2) the application owner is still a member of the Discord server, and
      3) that member still holds the Discord race role associated with the
         accepted character.

    This keeps old/archived applications from permanently consuming slots
    after a player leaves the server or loses their race role, while still
    using the Character Application to distinguish races such as Human vs
    Android (which intentionally share one Discord role).
    """

    counts = {
        race_tag: 0
        for race_tag in RACE_COUNT_LIMITS
    }

    inactive_departed = 0
    role_mismatch = 0

    for record in APPROVED_CHARACTER_RECORDS.values():
        race_tag_name = record.get("race")
        user_id = record.get("user_id")

        if race_tag_name not in counts or not user_id:
            continue

        member = guild.get_member(int(user_id))
        if member is None:
            inactive_departed += 1
            continue

        race_role_id = CHARACTER_RACE_ROLE_IDS.get(race_tag_name)
        if race_role_id is None:
            role_mismatch += 1
            continue

        member_role_ids = {
            role.id
            for role in member.roles
        }

        if race_role_id not in member_role_ids:
            role_mismatch += 1
            continue

        counts[race_tag_name] += 1

    return counts, inactive_departed, role_mismatch


def build_race_count_message(guild: discord.Guild):
    """Builds the public Race Count from the active Discord roster."""

    counts, _, _ = get_race_counts(guild)

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
        "-# Automatically updated by Golden Age Manager.",
        "-# Counts active Accepted characters only: the player must still be in",
        "-# the server and still hold the corresponding race role.",
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

        counts, departed_records, role_mismatches = get_race_counts(
            channel.guild
        )
        content = build_race_count_message(channel.guild)
        message = await get_race_count_message(channel)

        if reason == "startup" or departed_records or role_mismatches:
            active_total = sum(counts.values())
            diagnostic_text = ", ".join(
                f"{RACE_COUNT_LIMITS[race][0]}: {counts.get(race, 0)}"
                for race in RACE_COUNT_LIMITS
            )
            print(
                f"👥 Active race roster: {active_total} active accepted "
                f"character slot(s)."
            )
            print(f"📊 Active role-based breakdown: {diagnostic_text}")
            if departed_records:
                print(
                    f"🚪 Excluded {departed_records} Accepted character "
                    "record(s) because the player is no longer in the server."
                )
            if role_mismatches:
                print(
                    f"🏷️ Excluded {role_mismatches} Accepted character "
                    "record(s) because the player no longer has the matching "
                    "race role."
                )

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
            r"ability\s*to\s*mentor",
            r"mentor\s*ability",
            r"mentor\s*skill",
            r"skill",
        ],

        "reason": [
            # These aliases also act as boundaries so a compact field such as
            # "Ability to Mentor: Kaioken Ability Description: ..." does not
            # swallow the entire rest of the application into the skill name.
            r"ability\s*description",
            r"in[-\s]*character\s*justification(?:\s*&\s*alignment)?",
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
# MENTOR LIST / DIRECTORY
# ============================================================

async def rebuild_approved_mentor_records():
    """
    Rebuild the public mentor directory from the current approval tags on
    Mentor Applications (including archived threads).

    Only mentors whose Discord account is STILL a member of this server are
    considered active. Historical Mentor Applications remain untouched, but a
    departed/kicked member no longer appears in Mentor List and their mentor
    ability becomes available again.

    The source application is authoritative. The approved-mentors profile is
    a presentation copy and can be recreated if Staff deletes it by mistake.
    """

    APPROVED_MENTOR_RECORDS.clear()

    forum = await get_channel(MENTOR_APPLICATION_FORUM_ID)

    if not isinstance(forum, discord.ForumChannel):
        print(
            "⚠️ Could not rebuild mentor list: "
            "mentor-application forum not found."
        )
        return

    guild = forum.guild

    rejected_name = normalize(MENTOR_REJECTED_TAG)
    source_to_destination = {
        normalize(source): destination
        for source, destination in MENTOR_APPROVAL_MAPPING.items()
    }

    approved_threads_seen = 0
    skipped = 0
    departed_mentors = 0

    for thread in await collect_forum_threads(forum):
        tag_names = {
            normalize(tag.name)
            for tag in thread.applied_tags
        }

        if rejected_name in tag_names:
            continue

        mentorship_type = None
        for source_normalized, destination_name in source_to_destination.items():
            if source_normalized in tag_names:
                mentorship_type = destination_name
                break

        if mentorship_type is None:
            continue

        approved_threads_seen += 1

        starter_message = await get_starter_message(thread)
        if starter_message is None:
            skipped += 1
            print(
                f"⚠️ Approved mentor skipped: '{thread.name}' "
                "has no readable starter message."
            )
            continue

        # Active-roster validation:
        # Keep the historical application, but do not expose a mentor publicly
        # if the Discord account no longer exists in the guild.
        member = await get_guild_member(
            guild,
            starter_message.author.id
        )

        if member is None:
            departed_mentors += 1
            print(
                f"🚪 Inactive mentor excluded: '{thread.name}' "
                f"(Discord user {starter_message.author.id} is no longer "
                "in the server)."
            )
            continue

        application = parse_mentor_application(
            starter_message.content or ""
        )

        character = (
            application.get("character")
            or member.display_name
            or starter_message.author.display_name
            or thread.name
        )
        skill = application.get("skill") or thread.name

        APPROVED_MENTOR_RECORDS[thread.id] = {
            "user_id": starter_message.author.id,
            "author_name": member.display_name,
            "character": character,
            "skill": skill,
            "type": mentorship_type,
            "thread_name": thread.name,
        }

    print(
        f"🎓 Active Approved Mentor Applications counted: "
        f"{len(APPROVED_MENTOR_RECORDS)} mentor(s) "
        f"from {approved_threads_seen} approved thread(s)."
    )

    if departed_mentors:
        print(
            f"🚪 Excluded {departed_mentors} approved mentor application(s) "
            "because the player is no longer in the server."
        )

    if skipped:
        print(
            f"⚠️ Approved mentor applications skipped: {skipped}."
        )

    diagnostic_counts = {}
    for record in APPROVED_MENTOR_RECORDS.values():
        mentor_type = record.get("type", "Unknown")
        diagnostic_counts[mentor_type] = diagnostic_counts.get(mentor_type, 0) + 1

    if diagnostic_counts:
        diagnostic_text = ", ".join(
            f"{mentor_type}: {count}"
            for mentor_type, count in sorted(diagnostic_counts.items())
        )
        print(f"🎓 Active mentor breakdown: {diagnostic_text}")


def _clean_mentor_skill_name(skill: str, thread_name: str) -> str:
    """Return only the actual teachable ability/form/passive name.

    Mentor applications in the server come from several generations of forms.
    Older posts may use titles such as ``Mentor of Ki manipulation`` or values
    such as ``Kaioken (Single Form / Multiplier Skill)``. The public Mentor
    List and Available Abilities directory should show the ability itself, not
    form labels or application wording.
    """

    def clean_candidate(raw: str) -> str:
        value = str(raw or "").strip()
        if not value:
            return ""

        value = re.sub(r"[*_`~]", "", value).strip()

        normalized_value = normalize(value)
        bad_markers = (
            "ability description",
            "migration accepted",
            "in-character justification",
            "in character justification",
            "mentor reasoning",
            "mentor background",
            "lore reason",
        )
        if (
            len(value) > 120
            or "\n" in value
            or any(marker in normalized_value for marker in bad_markers)
        ):
            return ""

        value = re.sub(
            r"^(?:ability\s+to\s+mentor|skill\s+i\s+want\s+to\s+mentor|"
            r"mentor\s+ability|mentor\s+skill|ability|skill|technique|"
            r"custom\s+form|passive\s+skill)\s*[:\-–—]\s*",
            "",
            value,
            flags=re.IGNORECASE,
        ).strip()

        prefix_patterns = (
            r"^mentor\s+of\s+the\s+",
            r"^mentor\s+of\s+",
            r"^mentor\s+for\s+",
            r"^mentor\s*[:\-–—]\s*",
        )
        for prefix in prefix_patterns:
            value = re.sub(prefix, "", value, flags=re.IGNORECASE).strip()

        suffix_patterns = (
            r"\s+mentor$",
            r"\s+and\s+dojo\s+master(?:\s*\|.*)?$",
            r"\s*\|\s*[^|]+$",
        )
        for suffix in suffix_patterns:
            value = re.sub(suffix, "", value, flags=re.IGNORECASE).strip()

        value = re.sub(
            r"\s*\((?:single\s+form|multiplier\s+skill|single\s+form\s*/\s*"
            r"multiplier\s+skill|technique|custom\s+form|passive\s+skill|"
            r"mentor(?:ship)?)\)\s*$",
            "",
            value,
            flags=re.IGNORECASE,
        ).strip()

        value = re.sub(r"\s+", " ", value).strip(" -–—|:;")

        canonical_patterns = {
            r"^ki\s+manipulation$": "Ki Manipulation",
            r"^ki\s+protection$": "Ki Protection",
            r"^ki\s+amplification$": "Ki Amplification",
            r"^kaioken$": "Kaioken",
            r"^instant\s+transmission$": "Instant Transmission",
            r"^healing\s+reduction$": "Healing Reduction",
            r"^solar\s+flare$": "Solar Flare",
            r"^fusion$": "Fusion",
            r"^true\s+potential\s+unleashed$": "True Potential Unleashed",
        }
        for canonical_pattern, canonical_name in canonical_patterns.items():
            if re.fullmatch(canonical_pattern, value, flags=re.IGNORECASE):
                return canonical_name

        return value

    skill_value = clean_candidate(skill)
    thread_value = clean_candidate(thread_name)
    value = skill_value or thread_value or "Unknown Skill"

    if skill_value and thread_value:
        skill_norm = normalize(skill_value)
        thread_norm = normalize(thread_value)
        if (
            skill_norm == thread_norm
            or skill_norm.startswith(thread_norm + " ")
            or thread_norm.startswith(skill_norm + " ")
        ):
            value = min(
                (skill_value, thread_value),
                key=lambda item: (len(item), item.casefold()),
            )

    if _kienzan_bundle_status(value)[1]:
        return KIENZAN_BUNDLE_CATALOG_NAME

    if _death_beam_bundle_status(value)[1]:
        return DEATH_BEAM_BUNDLE_CATALOG_NAME

    return value or "Unknown Skill"


def _clean_mentor_character_name(character: str, author_name: str) -> str:
    """Keep mentor names compact even for older application formats."""

    value = str(character or "").strip()

    if (
        not value
        or len(value) > 80
        or "\n" in value
        or "migration" in normalize(value)
        or "ability to mentor" in normalize(value)
    ):
        value = str(author_name or "Unknown Mentor").strip()

    return value or "Unknown Mentor"


def build_mentor_list_message():
    """Build a very compact list: one mentor, followed by what they teach."""

    mentors = {}

    for record in APPROVED_MENTOR_RECORDS.values():
        character = _clean_mentor_character_name(
            record.get("character"),
            record.get("author_name"),
        )
        skill = _clean_mentor_skill_name(
            record.get("skill"),
            record.get("thread_name"),
        )

        key = character.casefold()
        entry = mentors.setdefault(
            key,
            {"character": character, "skills": []},
        )

        if skill.casefold() not in {item.casefold() for item in entry["skills"]}:
            entry["skills"].append(skill)

    ordered = sorted(
        mentors.values(),
        key=lambda item: item["character"].casefold(),
    )

    lines = [
        "## 🎓 Mentor List",
        "",
    ]

    if not ordered:
        lines.append("• No approved mentors yet.")
    else:
        for entry in ordered:
            skills = ", ".join(
                sorted(entry["skills"], key=str.casefold)
            )
            lines.append(
                f"• **{entry['character']}** — {skills}"
            )

    lines.extend([
        "",
        "-# Automatically updated by Golden Age Manager.",
    ])

    content = "\n".join(lines)

    if len(content) > 1990:
        content = content[:1900].rstrip() + (
            "\n\n-# Mentor List truncated because it exceeded Discord's message limit."
        )

    return content


async def get_mentor_list_message(channel: discord.TextChannel):
    """Find Golden Age Manager's managed Mentor List message."""

    meta = PROCESSED_APPLICATIONS.setdefault("meta", {})
    stored_message_id = meta.get("mentor_list_message_id")

    if stored_message_id:
        try:
            message = await channel.fetch_message(int(stored_message_id))
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
                and "## 🎓 Mentor List" in (message.content or "")
                and "Automatically updated by Golden Age Manager"
                in (message.content or "")
            ):
                meta["mentor_list_message_id"] = message.id
                save_processed_applications()
                return message
    except (discord.Forbidden, discord.HTTPException):
        pass

    return None


async def update_mentor_list_message(reason: str = "automatic update"):
    """Create or edit the managed Mentor List in #information."""

    async with MENTOR_LIST_LOCK:
        channel = await get_channel(MENTOR_LIST_CHANNEL_ID)

        if not isinstance(channel, discord.TextChannel):
            print(
                "⚠️ Mentor List channel was not found or is not a text channel."
            )
            return False

        content = build_mentor_list_message()
        message = await get_mentor_list_message(channel)

        try:
            if message is None:
                message = await channel.send(content)
                PROCESSED_APPLICATIONS.setdefault(
                    "meta", {}
                )["mentor_list_message_id"] = message.id
                save_processed_applications()
                print(
                    f"🎓 Mentor List message created in #{channel.name}."
                )
            elif message.content != content:
                await message.edit(content=content)
                print(
                    f"🎓 Mentor List updated ({reason})."
                )
            else:
                print(
                    f"🎓 Mentor List already current ({reason})."
                )

            return True

        except discord.Forbidden:
            print(
                "❌ Golden Age Manager cannot create/edit the Mentor List. "
                "Check View Channel, Send Messages and Read Message History."
            )
            return False

        except discord.HTTPException as error:
            print(
                f"❌ Discord API error while updating Mentor List: {error}"
            )
            return False



# ============================================================
# AVAILABLE ABILITIES DIRECTORY
# ============================================================

def _mentor_ability_match_key(value: str) -> str:
    """Normalize an ability name for catalogue/approval matching."""

    value = str(value or "").replace("_", " ").replace("-", " ")
    value = value.casefold()
    value = re.sub(r"[^a-z0-9']+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _kienzan_bundle_status(value: str):
    """Return (mentions_kienzan_family, has_required_pair).

    A valid Kienzan mentor application must explicitly cover BOTH the normal
    Kienzan/Destructo Disc and the double variant.
    """

    key = _mentor_ability_match_key(value)
    if not key:
        return False, False

    double_match = next(
        (name for name in KIENZAN_DOUBLE_NAMES if name in key),
        None,
    )

    remainder = key
    for name in KIENZAN_DOUBLE_NAMES:
        remainder = remainder.replace(name, " ")
    remainder = re.sub(r"\s+", " ", remainder).strip()

    single_match = any(
        re.search(rf"(?<![a-z0-9]){re.escape(name)}(?![a-z0-9])", remainder)
        for name in KIENZAN_SINGLE_NAMES
    )

    mentions_family = bool(double_match) or any(
        name in key for name in KIENZAN_SINGLE_NAMES
    )
    has_pair = bool(double_match) and bool(single_match)

    return mentions_family, has_pair


def _death_beam_bundle_status(value: str):
    """Return (mentions_death_beam_family, has_required_pair).

    A valid Death Beam mentor application must explicitly cover BOTH the normal
    Death Beam and Emperor Death Beam. Emperor Death Beam alone does not count
    as both just because its name contains the words "Death Beam".
    """

    key = _mentor_ability_match_key(value)
    if not key:
        return False, False

    emperor_match = next(
        (name for name in EMPEROR_DEATH_BEAM_NAMES if name in key),
        None,
    )

    # Remove the Emperor variant before checking for the standalone Death Beam.
    # This prevents "Emperor Death Beam" from satisfying both requirements.
    remainder = key
    for name in EMPEROR_DEATH_BEAM_NAMES:
        remainder = remainder.replace(name, " ")
    remainder = re.sub(r"\s+", " ", remainder).strip()

    single_match = any(
        re.search(rf"(?<![a-z0-9]){re.escape(name)}(?![a-z0-9])", remainder)
        for name in DEATH_BEAM_SINGLE_NAMES
    )

    mentions_family = bool(emperor_match) or any(
        re.search(rf"(?<![a-z0-9]){re.escape(name)}(?![a-z0-9])", key)
        for name in DEATH_BEAM_SINGLE_NAMES
    )
    has_pair = bool(emperor_match) and bool(single_match)

    return mentions_family, has_pair


def _mentor_bundle_requirement_error(skill: str, mentor_type: str):
    """Return an error message when a paired mentor slot is incomplete."""

    if mentor_type != "Technique":
        return None

    mentions_family, has_pair = _kienzan_bundle_status(skill)
    if mentions_family and not has_pair:
        return (
            "Kienzan and Kienzan Doble are a **single combined mentor slot**. "
            "The same mentor must know and apply for **both techniques**. "
            "Update the ability field to explicitly include "
            "`Kienzan + Kienzan Doble` before approving this application."
        )

    mentions_family, has_pair = _death_beam_bundle_status(skill)
    if mentions_family and not has_pair:
        return (
            "Death Beam and Emperor Death Beam are a **single combined mentor slot**. "
            "The same mentor must know and apply for **both techniques**. "
            "Update the ability field to explicitly include "
            "`Death Beam + Emperor Death Beam` before approving this application."
        )

    return None


def _canonical_catalog_ability(ability: str, mentor_type: str):
    """Return the exact catalogue name if an approved skill occupies a slot."""

    key = _mentor_ability_match_key(ability)
    if not key:
        return None

    # Paired techniques occupy one mentor slot and only count as occupied when
    # the approved application explicitly covers BOTH techniques in the pair.
    if mentor_type == "Technique":
        mentions_kienzan, has_kienzan_pair = _kienzan_bundle_status(ability)
        if has_kienzan_pair:
            return KIENZAN_BUNDLE_CATALOG_NAME
        if mentions_kienzan:
            return None

        mentions_death_beam, has_death_beam_pair = _death_beam_bundle_status(ability)
        if has_death_beam_pair:
            return DEATH_BEAM_BUNDLE_CATALOG_NAME
        if mentions_death_beam:
            return None

    # First resolve known aliases such as "Big Bang Attack" -> "Big Bang".
    alias_target = MENTOR_ABILITY_ALIASES.get(key)
    if alias_target and alias_target in MENTOR_ABILITY_CATALOG.get(mentor_type, []):
        return alias_target

    # Then compare directly against the configured catalogue names.
    for catalogue_name in MENTOR_ABILITY_CATALOG.get(mentor_type, []):
        if key == _mentor_ability_match_key(catalogue_name):
            return catalogue_name

    return None


def build_available_abilities_message():
    """Show only mentor abilities whose application slot is still OPEN.

    Formula:
        configured mentor catalogue - currently approved mentor abilities
        = abilities someone may still apply to mentor.

    Approved skills outside the configured catalogue are intentionally ignored;
    they belong in Mentor List but do not consume one of these whitelist slots.
    """

    occupied = {
        mentor_type: set()
        for mentor_type in MENTOR_ABILITY_CATALOG
    }

    for record in APPROVED_MENTOR_RECORDS.values():
        mentor_type = record.get("type")
        if mentor_type not in occupied:
            continue

        approved_ability = _clean_mentor_skill_name(
            record.get("skill"),
            record.get("thread_name"),
        )

        catalogue_name = _canonical_catalog_ability(
            approved_ability,
            mentor_type,
        )

        if catalogue_name:
            occupied[mentor_type].add(
                _mentor_ability_match_key(catalogue_name)
            )

    sections = [
        ("Technique", "### 🎯 Techniques"),
        ("Custom-Form", "### 🧬 Custom Forms"),
        ("Passive-Skill", "### ✨ Passive Skills"),
    ]

    lines = [
        "## 📚 Available Mentor Abilities",
        "",
        "These are the abilities that **currently have no approved mentor** and are open for Mentor Applications.",
        "",
        "-# Kienzan + Kienzan Doble share one mentor slot; the same mentor must know both techniques.",
        "-# Death Beam + Emperor Death Beam share one mentor slot; the same mentor must know both techniques.",
        "-# DragonMineZ 2.2 Ki, Strike and Evasion techniques are included in this catalogue.",
        "",
    ]

    configured_any = False
    available_total = 0

    for mentor_type, heading in sections:
        catalogue = MENTOR_ABILITY_CATALOG.get(mentor_type, [])

        # Staff has not supplied a catalogue for this category yet, so do not
        # pretend that approved abilities define what is available.
        if not catalogue:
            continue

        configured_any = True
        lines.append(heading)

        available = [
            ability
            for ability in catalogue
            if _mentor_ability_match_key(ability) not in occupied[mentor_type]
        ]
        available = sorted(available, key=str.casefold)
        available_total += len(available)

        if available:
            lines.extend(
                f"• {ability}"
                for ability in available
            )
        else:
            lines.append("• **No mentor slots currently available.**")

        lines.append("")

    if not configured_any:
        lines.append("• No mentor ability catalogue has been configured yet.")
    else:
        lines.append(
            f"-# {available_total} mentor slot(s) currently available. "
            "Automatically updated when Mentor Applications are approved, removed, or restored."
        )

    content = "\n".join(lines).rstrip()

    if len(content) > 1990:
        content = content[:1900].rstrip() + (
            "\n\n-# Available Mentor Abilities list truncated because it exceeded Discord's message limit."
        )

    return content


async def get_available_abilities_message(channel: discord.TextChannel):
    """Find Golden Age Manager's managed Available Abilities message."""

    meta = PROCESSED_APPLICATIONS.setdefault("meta", {})
    stored_message_id = meta.get("available_abilities_message_id")

    if stored_message_id:
        try:
            message = await channel.fetch_message(int(stored_message_id))
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
                and (
                    "## 📚 Available Mentor Abilities" in (message.content or "")
                    or "## 📚 Available Abilities" in (message.content or "")
                )
                and "Automatically updated" in (message.content or "")
            ):
                meta["available_abilities_message_id"] = message.id
                save_processed_applications()
                return message
    except (discord.Forbidden, discord.HTTPException):
        pass

    return None


async def update_available_abilities_message(reason: str = "automatic update"):
    """Create or edit the managed open Mentor Abilities catalogue in #information."""

    async with AVAILABLE_ABILITIES_LOCK:
        channel = await get_channel(AVAILABLE_ABILITIES_CHANNEL_ID)

        if not isinstance(channel, discord.TextChannel):
            print(
                "⚠️ Available Abilities channel was not found or is not a text channel."
            )
            return False

        content = build_available_abilities_message()
        message = await get_available_abilities_message(channel)

        try:
            if message is None:
                message = await channel.send(content)
                PROCESSED_APPLICATIONS.setdefault(
                    "meta", {}
                )["available_abilities_message_id"] = message.id
                save_processed_applications()
                print(
                    f"📚 Available Abilities message created in #{channel.name}."
                )
            elif message.content != content:
                await message.edit(content=content)
                print(
                    f"📚 Available Abilities updated ({reason})."
                )
            else:
                print(
                    f"📚 Available Abilities already current ({reason})."
                )

            return True

        except discord.Forbidden:
            print(
                "❌ Golden Age Manager cannot create/edit Available Abilities. "
                "Check View Channel, Send Messages and Read Message History."
            )
            return False

        except discord.HTTPException as error:
            print(
                f"❌ Discord API error while updating Available Abilities: {error}"
            )
            return False


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
            rf"(?=\s+(?:birth\s*name|character\s*name|name|born|designation|affiliation|alignment|class|"
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

    # The explicit labeled Race/Species field is authoritative.
    # Generic parsing can swallow following fields in compact/legacy posts,
    # so always prefer the dedicated extractor when it succeeds.
    explicit_race = extract_character_race_fallback(content)

    if explicit_race:
        previous_race = application.get("race")
        application["race"] = explicit_race

        if not previous_race:
            print(
                f"🧬 Race parser fallback detected: "
                f"{explicit_race!r}"
            )
        elif normalize(previous_race) != normalize(explicit_race):
            print(
                f"🧬 Race parser corrected labeled value: "
                f"{previous_race!r} -> {explicit_race!r}"
            )

    return application


# ============================================================
# RACE DETECTION
# ============================================================

def detect_race_tag(
    race_value: str
):
    """
    Resolve the race written in the explicit Race/Species field.

    v7.4 keeps the v7.3 protection against backstory mentions and adds a
    dedicated Human/Android disambiguation pass. Discord applications have
    historically used several equivalent spellings such as ``Human/Androids``,
    ``Human - Android`` and ``Human (Android)``. Those must resolve to the
    Android slot before the generic ``Human`` alias is considered.
    """

    if not race_value:
        return None

    race_normalized = normalize(race_value)
    if not race_normalized:
        return None

    # Only inspect the beginning of the labeled value for special compound
    # races. This prevents words later in a backstory from changing the race.
    leading = race_normalized[:80]
    leading_words = re.sub(r"[^a-z0-9]+", " ", leading).strip()
    leading_words = re.sub(r"\s+", " ", leading_words)

    # Human/Android is a distinct race bucket. Check it BEFORE plain Human.
    # Accept common historical variants: Human/Android(s), Human-Android,
    # Human (Android), Android/Human, etc.
    if re.match(r"^human\s+androids?\b", leading_words):
        return "Android"
    if re.match(r"^androids?\s+human\b", leading_words):
        return "Android"

    # Human-derived cyborg / organic-android descriptions still occupy the
    # Human (Android) bucket. Examples:
    #   Human (Cyborg / Organic Android)
    #   Human (Cyborg)
    #   Human / Organic Android
    if leading_words.startswith("human ") and re.search(
        r"\b(?:androids?|cyborgs?|organic\s+androids?|artificial\s+humans?)\b",
        leading_words,
    ):
        return "Android"

    # Bio-Android must also win before the generic Android alias.
    if re.match(r"^bio\s+androids?\b", leading_words):
        return "Bio-Android"

    # v8.0 — New races. Check common explicit forms before generic alias
    # matching so punctuation/possessives are handled consistently.
    if re.match(r"^(?:grey|gray)(?:\s+race|\s+alien)?\b", leading_words):
        return "Grey Race"
    if re.match(r"^jiren(?:s)?(?:\s+race|\s+species)?\b", leading_words):
        return "Grey Race"

    if re.match(r"^demon(?:s|ic)?(?:\s+race|\s+clan|\s+realm(?:\s+race)?)?\b", leading_words):
        return "Demon Race"

    # 1) Exact match or explicit race at the beginning of the field. Longest
    # aliases are checked first so compound races win over generic aliases.
    prefix_matches = []

    for written_race, target_tag in RACE_MAPPING.items():
        written_normalized = normalize(written_race)
        if not written_normalized:
            continue

        if race_normalized == written_normalized:
            prefix_matches.append((len(written_normalized), target_tag))
            continue

        if race_normalized.startswith(written_normalized):
            remainder = race_normalized[len(written_normalized):]
            if not remainder or not remainder[0].isalnum():
                prefix_matches.append((len(written_normalized), target_tag))
            elif remainder[0].isspace():
                prefix_matches.append((len(written_normalized), target_tag))

    if prefix_matches:
        prefix_matches.sort(key=lambda item: item[0], reverse=True)
        return prefix_matches[0][1]

    # 2) Legacy compatibility only for compact race values. Never scan a
    # paragraph/backstory for race words.
    if len(race_normalized) <= 80:
        matches = []
        for written_race, target_tag in RACE_MAPPING.items():
            written_normalized = normalize(written_race)
            if written_normalized and written_normalized in race_normalized:
                matches.append((len(written_normalized), target_tag))

        if matches:
            matches.sort(key=lambda item: item[0], reverse=True)
            return matches[0][1]

    return None


# ============================================================
# FORM APPLICATION PARSER
# ============================================================

def parse_form_application(content: str):
    """Parse the public Form Application template used by #form-application."""

    aliases = {
        "character_name": [
            r"character\s*name",
            r"character",
        ],
        "current_level": [
            r"current\s*level",
            r"level",
        ],
        "time_played": [
            r"time\s*played",
            r"play\s*time",
            r"time\s*on\s*server",
        ],
        "form_requested": [
            r"form\s*requested",
            r"requested\s*form",
            r"transformation\s*requested",
            r"form",
        ],
        "reasoning": [
            r"reasoning",
            r"reason",
            r"justification",
        ],
        "roleplay_plan": [
            r"roleplay\s*plan",
            r"rp\s*plan",
            r"unlock\s*plan",
        ],
    }

    return parse_labeled_fields(content, aliases)


# ============================================================
# STAFF DECLINE REASON / FEEDBACK SYSTEM
# ============================================================

DECLINE_REASON_PATTERN = re.compile(
    r"(?:^|\n)\s*declined\s*:\s*(.+)$",
    re.IGNORECASE | re.DOTALL,
)


def _member_can_supply_decline_reason(member: discord.Member) -> bool:
    """Only Staff/moderation-capable members may supply an official reason."""

    if member.guild.owner_id == member.id:
        return True

    permissions = member.guild_permissions
    return any((
        permissions.administrator,
        permissions.manage_guild,
        permissions.manage_threads,
        permissions.manage_messages,
        permissions.moderate_members,
    ))


async def find_latest_staff_decline_reason(thread: discord.Thread):
    """
    Find the newest Staff message in the application thread containing:

        Declined: <reason>

    The applicant cannot spoof the reason because only moderation-capable
    members are considered authoritative.
    """

    try:
        async for message in thread.history(limit=100, oldest_first=False):
            if message.author.bot:
                continue

            if not isinstance(message.author, discord.Member):
                continue

            if not _member_can_supply_decline_reason(message.author):
                continue

            content = (message.content or "").strip()
            if not content:
                continue

            match = DECLINE_REASON_PATTERN.search(content)
            if not match:
                continue

            reason = clean_value(match.group(1) or "")
            if not reason:
                continue

            # Keep Discord messages safely below the 2000-character limit.
            reason = reason[:1400].strip()
            return reason, message.author, message

    except discord.Forbidden:
        print(
            "❌ Golden Age Manager cannot read application history to find "
            "the Declined: reason. Check Read Message History."
        )
    except discord.HTTPException as error:
        print(f"❌ Discord API error while reading decline reason: {error}")

    return None


def build_decline_suggestions(reason: str, application_kind: str) -> list[str]:
    """
    Rule-based contextual suggestions derived from the Staff reason.

    This intentionally does not overrule Staff or invent a new rejection
    reason. It turns common feedback themes into practical revision ideas.
    """

    value = normalize(reason)
    kind = normalize(application_kind)
    suggestions = []

    def add(*items):
        for item in items:
            if item not in suggestions:
                suggestions.append(item)

    # Capacity / race cap
    if any(term in value for term in (
        "race cap", "cap reached", "race full", "slots full",
        "no slots", "full race", "capacity", "slot unavailable",
    )):
        add(
            "Check the current Race Count before resubmitting.",
            "Consider adapting the concept to a race that currently has an open slot.",
            "If you want to keep this race, wait until a slot becomes available before reapplying.",
        )

    # Powerplaying / guaranteed exceptional status
    if any(term in value for term in (
        "powerplay", "power playing", "overpowered", "too powerful",
        "godlike", "god-like", "guaranteed", "legendary",
        "unknown potential", "unlimited", "perfect being",
        "no limits", "future power", "too strong",
    )):
        add(
            "Remove guaranteed future power, transformations, or special status from the concept.",
            "Give the character meaningful weaknesses, consequences, and situations where they can fail.",
            "Move major growth into future roleplay instead of establishing it as guaranteed in the application.",
        )

    # Weakness / balance concerns
    if any(term in value for term in (
        "weakness", "weaknesses", "limitation", "limitations",
        "drawback", "drawbacks", "balance", "balanced", "tradeoff",
    )):
        add(
            "Make the weaknesses concrete and usable in actual roleplay, not only cosmetic.",
            "Explain what the character cannot do and what consequences come with their strengths.",
            "Add clear tradeoffs so strong abilities create meaningful risks or costs.",
        )

    # Lore fit / origin
    if any(term in value for term in (
        "lore", "canon", "setting", "origin", "backstory",
        "doesn't fit", "does not fit", "incompatible", "timeline",
    )):
        add(
            "Adjust the origin and backstory so they fit the server's established lore and timeline.",
            "Connect important claims to events, factions, locations, or circumstances that can exist in the setting.",
            "Avoid relying on exceptional canon-level events unless Staff has explicitly approved them.",
        )

    # Detail / clarity / writing quality
    if any(term in value for term in (
        "detail", "details", "too short", "unclear", "vague",
        "not enough", "insufficient", "underdeveloped", "poorly explained",
        "needs explanation", "lack of explanation",
    )):
        add(
            "Expand the parts Staff identified as unclear instead of only adding more general text.",
            "Use concrete examples of how the concept will work during roleplay.",
            "Make motivations, limitations, and intended development explicit.",
        )

    # Conduct / RP history / activity
    if any(term in value for term in (
        "conduct", "behavior", "behaviour", "roleplay history", "rp history",
        "activity", "inactive", "time played", "new player", "unestablished",
    )):
        add(
            "Spend more time establishing the character through normal roleplay before reapplying.",
            "Demonstrate consistent rule-following, activity, and collaborative roleplay.",
            "Use future RP scenes to show the qualities Staff needs to evaluate rather than only describing them in text.",
        )

    # Mentor-specific themes
    if kind == "mentor":
        if any(term in value for term in (
            "experience", "mastery", "doesn't know", "does not know",
            "knowledge", "qualified", "qualification", "training",
        )):
            add(
                "Show clear in-character experience and mastery of the technique before applying again.",
                "Add training, battles, or lore events that demonstrate why this character is qualified to teach it.",
                "Explain how the character would actually teach the skill to other players.",
            )
        if any(term in value for term in ("availability", "inactive", "activity")):
            add(
                "Make sure the mentor can be consistently available to train other players.",
            )

    # Form-specific themes
    if kind == "form":
        if any(term in value for term in (
            "reason", "reasoning", "weak reasoning", "motivation",
            "doesn't justify", "does not justify", "earned", "earn",
        )):
            add(
                "Tie the form to the character's existing story instead of only wanting the transformation for power.",
                "Describe the RP circumstances that would make the unlock feel earned.",
                "Strengthen the connection between the form, the character's motivations, and recent development.",
            )
        if any(term in value for term in (
            "not open", "not announced", "closed", "unavailable form",
            "not available", "too early",
        )):
            add(
                "Wait until Staff officially announces that this form is open for applications.",
                "Do not resubmit the same form while its application window is closed.",
            )

    # Character-specific fallback tailoring
    if kind == "character" and any(term in value for term in (
        "character", "concept", "personality", "motivation",
    )):
        add(
            "Revise the character concept directly around the concern Staff identified.",
            "Make sure strengths, flaws, motivations, and future development create opportunities for other players to interact with.",
        )

    # Universal fallback: always give something useful without inventing facts.
    add(
        "Revise the application specifically around the Staff reason above before resubmitting.",
        "Make the changes explicit in the application so Staff can quickly see what was corrected.",
    )

    return suggestions[:4]


async def remove_application_status_tag(
    thread: discord.Thread,
    tag_name: str,
    normalizer=normalize,
):
    """Remove one status tag while preserving every other forum tag."""

    target = normalizer(tag_name)
    remaining_tags = [
        tag
        for tag in thread.applied_tags
        if normalizer(tag.name) != target
    ]

    await thread.edit(
        applied_tags=remaining_tags,
        reason="Golden Age Manager: missing required Declined reason",
    )


async def require_decline_reason_and_reject(
    thread: discord.Thread,
    application_label: str,
    application_kind: str,
    rejection_tag_name: str,
    status_word: str = "Rejected",
    normalizer=normalize,
) -> bool:
    """
    Enforce Staff's `Declined: <reason>` format before closing an application.

    Returns True when the application was actually closed, False when the
    status tag had to be removed because Staff did not provide a reason.
    """

    result = await find_latest_staff_decline_reason(thread)

    if result is None:
        try:
            await remove_application_status_tag(
                thread,
                rejection_tag_name,
                normalizer=normalizer,
            )

            await thread.send(
                "## ⚠️ Missing Decline Reason\n\n"
                "Before closing this application, a Staff member must post an "
                "official reason in this thread using:\n\n"
                "`Declined: <reason>`\n\n"
                f"The **{status_word}** tag was removed automatically. "
                "The application remains open."
            )

            print(
                f"⚠️ {application_label} close blocked: no Staff "
                "Declined: reason was found."
            )

        except discord.Forbidden:
            print(
                "❌ Could not remove the status tag after a missing decline "
                "reason. Check Manage Threads / Send Messages."
            )
        except discord.HTTPException as error:
            print(f"❌ Discord API error enforcing decline reason: {error}")

        return False

    reason, staff_member, _ = result
    suggestions = build_decline_suggestions(reason, application_kind)

    await reject_application(
        thread,
        application_label,
        reason=reason,
        suggestions=suggestions,
        staff_member=staff_member,
        status_word=status_word,
    )

    return True


# ============================================================
# REJECTION AUTOMATION
# ============================================================

async def reject_application(
    thread: discord.Thread,
    application_label: str,
    reason: str | None = None,
    suggestions: list[str] | None = None,
    staff_member: discord.Member | None = None,
    status_word: str = "Rejected",
):
    """Close an application while preserving Staff's written reasoning."""

    try:
        heading_emoji = "🚫" if normalize(status_word) == "voided" else "❌"

        parts = [
            f"## {heading_emoji} {application_label} {status_word}",
            "",
            f"Your {application_label.lower()} has been "
            f"{status_word.lower()} by the Staff Team.",
        ]

        if reason:
            parts.extend([
                "",
                "### 📋 Staff Reason",
                f"> {reason.replace(chr(10), chr(10) + '> ')}",
            ])

        if suggestions:
            parts.extend([
                "",
                "### 💡 Golden Age Suggestions",
            ])
            parts.extend(
                f"• {suggestion}"
                for suggestion in suggestions
            )
            parts.extend([
                "",
                "These suggestions are based on the Staff reason above and "
                "are meant to help with a future revision. Staff's decision "
                "and reasoning remain authoritative.",
            ])

        if staff_member is not None:
            parts.extend([
                "",
                f"**Reviewed by:** {staff_member.mention}",
            ])

        await thread.send("\n".join(parts)[:1990])

        await thread.edit(
            archived=True,
            locked=True,
            reason=(
                "Golden Age Manager: "
                f"{application_label} {status_word.lower()}"
            )
        )

        print(
            f"❌ {application_label} {status_word.lower()} with Staff reason."
        )
        print("🔒 Application archived and locked.")

    except discord.Forbidden:
        print(
            "❌ Golden Age Manager doesn't have permission to close "
            "this application."
        )

    except discord.HTTPException as error:
        print(
            f"❌ Discord API error while closing application: {error}"
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

        "form-application":
            FORM_APPLICATION_FORUM_ID,

        "approved-forms":
            APPROVED_FORMS_FORUM_ID,
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
        # First purge stale managed publications left behind while the bot was
        # offline. This also removes their bot-generated approved mirrors.
        for guild in client.guilds:
            await cleanup_departed_member_publications(
                guild,
                reason="startup roster audit",
            )

        await rebuild_approved_mentor_records()
        await update_mentor_list_message(reason="startup")
        await update_available_abilities_message(reason="startup")
        await rebuild_approved_character_cache()
        # Mutant applications are retired. Race Count now reconciles Accepted
        # Character Applications against current Discord membership + race roles.
        await update_race_count_message(reason="startup")
        print("🧬 Legacy Mutant application/count system: RETIRED (v7.5).")
        print("📝 Declined: feedback system enabled for Character, Mentor, and Form Applications.")
        print("🔥 Form Application automation enabled (v7.5).")
        print("👥 Active Race Count reconciliation enabled (v7.5).")
        print("🧬 Explicit Race/Species field priority fix enabled (v7.5).")
        print("🤖 Human/Android race disambiguation enabled (v7.5).")
        print("🟠 Saiyan race cap set to 11 (v7.5).")
        print("🧠 Robust compact-field race parsing enabled (v7.5).")
        print("🎓 Active Mentor roster reconciliation enabled (v8.1).")
        print("🧹 Departed-member publication cleanup enabled (v8.1).")
        print("👽 Grey Race + 😈 Demon Race support enabled (v8.1, cap 3 each).")
        print("🥋 DragonMineZ 2.2 Mentor Technique catalogue enabled (v8.1).")
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

    print(
        "📚 Per-character Story ZIP command: !storyreport"
    )

    print(
        "📖 Consolidated Story Word command: !storyreportall"
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
    """
    Keeps Community Access aligned with Member + Verified and refreshes the
    public Race Count whenever a relevant race/access role changes.
    """

    access_role_ids = {
        MEMBER_ROLE_ID,
        VERIFIED_ROLE_ID,
        COMMUNITY_ACCESS_ROLE_ID,
    }

    race_role_ids = set(
        CHARACTER_RACE_ROLE_IDS.values()
    )

    before_access_ids = {
        role.id
        for role in before.roles
        if role.id in access_role_ids
    }

    after_access_ids = {
        role.id
        for role in after.roles
        if role.id in access_role_ids
    }

    before_race_ids = {
        role.id
        for role in before.roles
        if role.id in race_role_ids
    }

    after_race_ids = {
        role.id
        for role in after.roles
        if role.id in race_role_ids
    }

    access_changed = (
        before_access_ids != after_access_ids
    )

    race_changed = (
        before_race_ids != after_race_ids
    )

    if not access_changed and not race_changed:
        return

    if access_changed:
        sync_key = (after.guild.id, after.id)

        if sync_key not in ACCESS_SYNC_IN_PROGRESS:
            await sync_community_access_for_member(
                after,
                reason=(
                    "Golden Age Manager: Member + Verified access gate"
                )
            )

    if race_changed:
        await update_race_count_message(
            reason=f"race role change for {after}"
        )


@client.event
async def on_member_remove(member: discord.Member):
    """
    Purge managed publications when a player leaves/is removed, then rebuild
    every public directory/count that depends on those applications.

    This event is fired for normal leaves, kicks and bans as a member removal.
    """

    if member.bot:
        return

    await cleanup_departed_member_publications(
        member.guild,
        user_id=member.id,
        reason=f"member left: {member}",
    )

    # Character source applications may have been deleted, so rebuild the
    # accepted-character cache before recalculating Race Count.
    await rebuild_approved_character_cache()
    await update_race_count_message(
        reason=f"member left: {member}"
    )

    # Mentor applications/mirrors may also have been deleted. Rebuilding here
    # automatically frees their abilities in Available Mentor Abilities.
    await rebuild_approved_mentor_records()
    await update_mentor_list_message(
        reason=f"member left: {member}"
    )
    await update_available_abilities_message(
        reason=f"member left: {member}"
    )

    print(
        f"🚪 Member left the server: {member}. "
        "Managed publications deleted and public registries reconciled."
    )


@client.event
async def on_member_join(member: discord.Member):
    """
    Reconcile mentors when a member joins/rejoins.

    If an old Mentor Application is still approved, the mentor can return to
    the public Mentor List automatically after rejoining the server.
    """

    if member.bot:
        return

    await rebuild_approved_mentor_records()
    await update_mentor_list_message(
        reason=f"member joined: {member}"
    )
    await update_available_abilities_message(
        reason=f"member joined: {member}"
    )

    print(
        f"👋 Member joined the server: {member}. "
        "Mentor List + Available Abilities reconciled."
    )


# ============================================================
# STAFF REPORT COMMAND
# ============================================================

def _word_safe_text(value) -> str:
    """Remove control characters that cannot be stored in a DOCX file."""

    if value is None:
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return "".join(
        character
        for character in text
        if character in ("\n", "\t") or ord(character) >= 32
    ).strip()


def _application_status_label(thread: discord.Thread) -> str:
    """Return a readable status for one Character Application thread."""

    tag_names = {
        normalize_forum_tag(tag.name)
        for tag in thread.applied_tags
    }

    if normalize_forum_tag(CHARACTER_ACCEPTED_TAG) in tag_names:
        return "Accepted"

    if normalize_forum_tag(CHARACTER_REJECTED_TAG) in tag_names:
        return "Rejected"

    return "Pending"


async def collect_character_applications_for_story_report():
    """Index every Character Application by its Discord author ID."""

    forum = await get_channel(CHARACTER_APPLICATION_FORUM_ID)
    if not isinstance(forum, discord.ForumChannel):
        raise RuntimeError("The Character Application forum could not be found.")

    applications_by_author = {}

    for thread in await collect_forum_threads(forum):
        starter = await get_starter_message(thread)
        if starter is None:
            continue

        parsed = parse_character_application(starter.content or "")
        record = {
            "thread_name": thread.name,
            "url": starter.jump_url,
            "status": _application_status_label(thread),
            "character_name": clean_value(parsed.get("character_name") or ""),
            "race": clean_value(parsed.get("race") or ""),
            "application": parsed,
        }
        applications_by_author.setdefault(starter.author.id, []).append(record)

    return applications_by_author


async def collect_story_entries(channel):
    """Collect story posts from a forum, or messages from a text channel."""

    entries = []

    if isinstance(channel, discord.ForumChannel):
        threads = await collect_forum_threads(channel)
        threads.sort(key=lambda item: item.created_at or datetime.min.replace(tzinfo=ZoneInfo("UTC")))

        for thread in threads:
            starter = await get_starter_message(thread)
            if starter is None:
                continue

            messages = []
            try:
                async for item in thread.history(limit=None, oldest_first=True):
                    # A story thread may contain Staff feedback or comments
                    # from other players. Only the original author's posts
                    # belong in that character's narrative dossier.
                    if item.author.bot or item.author.id != starter.author.id:
                        continue
                    body = _word_safe_text(item.content)
                    attachments = [attachment.url for attachment in item.attachments]
                    if body or attachments:
                        messages.append({
                            "author": item.author.display_name,
                            "author_id": item.author.id,
                            "created_at": item.created_at,
                            "content": body,
                            "attachments": attachments,
                            "url": item.jump_url,
                        })
            except (discord.Forbidden, discord.HTTPException) as error:
                print(f"⚠️ Could not read story thread '{thread.name}': {error}")

            entries.append({
                "title": thread.name,
                "author": starter.author,
                "created_at": starter.created_at,
                "url": starter.jump_url,
                "messages": messages,
            })

        return entries

    if isinstance(channel, discord.TextChannel):
        async for item in channel.history(limit=None, oldest_first=True):
            if item.author.bot:
                continue
            body = _word_safe_text(item.content)
            attachments = [attachment.url for attachment in item.attachments]
            if not body and not attachments:
                continue
            entries.append({
                "title": f"Story by {item.author.display_name}",
                "author": item.author,
                "created_at": item.created_at,
                "url": item.jump_url,
                "messages": [{
                    "author": item.author.display_name,
                    "author_id": item.author.id,
                    "created_at": item.created_at,
                    "content": body,
                    "attachments": attachments,
                    "url": item.jump_url,
                }],
            })
        return entries

    raise RuntimeError("The configured story channel is not a Forum or Text channel.")


def get_story_report_stats(entries, applications_by_author):
    """Return story/post and player coverage metrics for the report."""

    player_ids = {
        entry["author"].id
        for entry in entries
        if entry.get("author") is not None
    }

    accepted_application_keys = set()

    for user_id in player_ids:
        for app in applications_by_author.get(user_id, []):
            if app.get("status") != "Accepted":
                continue

            # Application URLs are unique per Discord thread. Keep a fallback
            # key for safety in case an old record does not have a URL.
            application_key = (
                app.get("url")
                or f"{user_id}:{app.get('thread_name', '')}:{app.get('character_name', '')}"
            )
            accepted_application_keys.add(application_key)

    return {
        "stories": len(entries),
        "unique_players": len(player_ids),
        "accepted_characters": len(accepted_application_keys),
    }


def build_story_report_document(entries, applications_by_author, output_path):
    """Build the consolidated Character Stories Word document."""

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.color.rgb = RGBColor(76, 45, 122)

    document.add_heading("Dragon Ball: Golden Age", 0)
    document.add_heading("Character Stories Report", level=1)
    generated = datetime.now(ZoneInfo(REPORT_TIMEZONE))
    stats = get_story_report_stats(entries, applications_by_author)
    document.add_paragraph(
        f"Generated: {generated.strftime('%Y-%m-%d %H:%M %Z')}\n"
        f"Stories collected: {stats['stories']}\n"
        f"Unique players represented: {stats['unique_players']}\n"
        f"Accepted characters represented: {stats['accepted_characters']}"
    )

    note = document.add_paragraph(
        "Note: Stories are individual posts/threads. One player may have multiple stories."
    )
    note.runs[0].italic = True

    document.add_heading("Index", level=1)
    index_table = document.add_table(rows=1, cols=4)
    index_table.style = "Table Grid"
    for cell, label in zip(
        index_table.rows[0].cells,
        ("Story", "Player", "Application", "Character / Race"),
    ):
        cell.text = label

    for entry in entries:
        linked_apps = applications_by_author.get(entry["author"].id, [])
        application_text = ", ".join(app["status"] for app in linked_apps) or "Not found"
        character_text = "; ".join(
            " / ".join(filter(None, (app["character_name"], app["race"])))
            for app in linked_apps
        ) or "—"
        cells = index_table.add_row().cells
        cells[0].text = _word_safe_text(entry["title"])
        cells[1].text = _word_safe_text(entry["author"].display_name)
        cells[2].text = application_text
        cells[3].text = character_text

    for position, entry in enumerate(entries):
        document.add_page_break()
        document.add_heading(_word_safe_text(entry["title"]) or "Untitled Story", level=1)

        metadata = document.add_table(rows=0, cols=2)
        metadata.style = "Light Shading Accent 1"
        metadata_rows = (
            ("Player", entry["author"].display_name),
            ("Discord ID", str(entry["author"].id)),
            ("Published", entry["created_at"].astimezone(ZoneInfo(REPORT_TIMEZONE)).strftime("%Y-%m-%d %H:%M %Z")),
            ("Discord link", entry["url"]),
        )
        for label, value in metadata_rows:
            cells = metadata.add_row().cells
            cells[0].text = label
            cells[1].text = _word_safe_text(value)

        linked_apps = applications_by_author.get(entry["author"].id, [])
        document.add_heading("Character Application", level=2)
        if not linked_apps:
            document.add_paragraph("No Character Application was found for this Discord user.")
        else:
            for app in linked_apps:
                table = document.add_table(rows=0, cols=2)
                table.style = "Table Grid"
                for label, value in (
                    ("Character", app["character_name"] or app["thread_name"]),
                    ("Race", app["race"] or "Not detected"),
                    ("Status", app["status"]),
                    ("Application link", app["url"]),
                ):
                    cells = table.add_row().cells
                    cells[0].text = label
                    cells[1].text = _word_safe_text(value)

        document.add_heading("Story / Lore", level=2)
        if not entry["messages"]:
            document.add_paragraph("No readable text or attachments were found.")

        for message_index, item in enumerate(entry["messages"]):
            if message_index > 0:
                paragraph = document.add_paragraph()
                run = paragraph.add_run(
                    f"Reply by {item['author']} — "
                    f"{item['created_at'].astimezone(ZoneInfo(REPORT_TIMEZONE)).strftime('%Y-%m-%d %H:%M')}"
                )
                run.bold = True

            if item["content"]:
                for block in item["content"].split("\n"):
                    document.add_paragraph(block)

            if item["attachments"]:
                attachment_paragraph = document.add_paragraph()
                attachment_paragraph.add_run("Attachments:\n").bold = True
                for attachment_url in item["attachments"]:
                    attachment_paragraph.add_run(f"{attachment_url}\n")

    document.save(output_path)


def _matching_text(value) -> str:
    """Normalize names and thread titles for conservative matching."""

    value = unicodedata.normalize("NFKD", _word_safe_text(value).casefold())
    value = "".join(character for character in value if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def _safe_report_filename(value: str, fallback: str = "character") -> str:
    """Return a Windows/ZIP-safe filename component."""

    value = unicodedata.normalize("NFKD", _word_safe_text(value))
    value = "".join(character for character in value if not unicodedata.combining(character))
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", value)
    value = re.sub(r"\s+", "_", value.strip())
    value = re.sub(r"_+", "_", value).strip("._ ")
    return (value or fallback)[:80]


def choose_story_application(entry, applications_by_author):
    """
    Match one story to one Character Application without guessing.

    Accepted applications are preferred. A single candidate is safe to use.
    When a player owns several characters, the character name must appear in
    the story thread title. Ambiguous stories are returned as unresolved.
    """

    linked_apps = list(applications_by_author.get(entry["author"].id, []))
    if not linked_apps:
        return None

    accepted_apps = [app for app in linked_apps if app.get("status") == "Accepted"]
    candidates = accepted_apps or linked_apps

    if len(candidates) == 1:
        return candidates[0]

    title = _matching_text(entry.get("title") or "")
    matches = []

    for app in candidates:
        character_name = _matching_text(
            app.get("character_name") or app.get("thread_name") or ""
        )
        if character_name and character_name in title:
            matches.append((len(character_name), app))

    if not matches:
        return None

    # Prefer the longest matching name (for example, "Slug Jr" over "Slug").
    matches.sort(key=lambda item: item[0], reverse=True)
    if len(matches) == 1 or matches[0][0] > matches[1][0]:
        return matches[0][1]

    return None


def group_stories_by_character(entries, applications_by_author):
    """Group story threads into resolved character dossiers."""

    groups = {}

    for entry in entries:
        application = choose_story_application(entry, applications_by_author)

        if application is not None:
            application_key = (
                application.get("url")
                or f"{entry['author'].id}:{application.get('thread_name', '')}"
            )
            key = ("character", application_key)
            character_name = (
                application.get("character_name")
                or application.get("thread_name")
                or entry["title"]
            )
            race = application.get("race") or "Unknown Race"
            unresolved = False
        else:
            # All ambiguous stories belonging to the same player are kept
            # together for manual review instead of being assigned at random.
            key = ("unresolved", entry["author"].id)
            character_name = entry["author"].display_name
            race = "Unknown character"
            unresolved = True

        group = groups.setdefault(key, {
            "character_name": character_name,
            "race": race,
            "author": entry["author"],
            "application": application,
            "entries": [],
            "unresolved": unresolved,
        })
        group["entries"].append(entry)

    return list(groups.values())


def _add_compact_label(document, label: str, value: str):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(1)
    label_run = paragraph.add_run(f"{label}: ")
    label_run.bold = True
    paragraph.add_run(_word_safe_text(value) or "—")
    return paragraph


def build_character_story_document(group, output_path):
    """Build one compact Word dossier for one character."""

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    styles = document.styles
    normal_style = styles["Normal"]
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(9)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(2)
    normal_style.paragraph_format.line_spacing = 1

    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.size = Pt(22)
    styles["Title"].font.color.rgb = RGBColor(76, 45, 122)

    for heading_name, size in (("Heading 1", 15), ("Heading 2", 12)):
        styles[heading_name].font.name = "Aptos Display"
        styles[heading_name].font.size = Pt(size)
        styles[heading_name].paragraph_format.space_before = Pt(7)
        styles[heading_name].paragraph_format.space_after = Pt(3)

    document_title = (
        f"Unresolved Stories — {group['author'].display_name}"
        if group["unresolved"]
        else group["character_name"]
    )
    document.add_heading(_word_safe_text(document_title), 0)
    subtitle = document.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(5)
    subtitle_run = subtitle.add_run(
        f"{_word_safe_text(group['race'])} • {_word_safe_text(group['author'].display_name)}"
    )
    subtitle_run.bold = True

    _add_compact_label(document, "Discord ID", str(group["author"].id))
    _add_compact_label(document, "Stories included", str(len(group["entries"])))

    application = group.get("application")
    document.add_heading("Character Application", level=1)

    if application is None:
        warning = document.add_paragraph(
            "This player's story could not be matched safely to one Character Application. "
            "Review the Discord links manually before using this dossier for evaluation."
        )
        if warning.runs:
            warning.runs[0].bold = True
    else:
        _add_compact_label(document, "Status", application.get("status") or "Unknown")
        _add_compact_label(document, "Application link", application.get("url") or "")

        parsed = application.get("application") or {}
        application_fields = (
            ("Alias", parsed.get("alias")),
            ("Class", parsed.get("class")),
            ("Age", parsed.get("age")),
            ("Sex", parsed.get("sex")),
            ("Family", parsed.get("family")),
            ("Alignment", parsed.get("alignment")),
            ("Fighting style", parsed.get("fighting_style")),
        )
        for label, value in application_fields:
            if value:
                _add_compact_label(document, label, value)

        application_backstory = _word_safe_text(parsed.get("backstory") or "")
        if application_backstory:
            document.add_heading("Application Backstory", level=2)
            paragraph = document.add_paragraph()
            paragraph.add_run(application_backstory)

    document.add_heading("Stories / Roleplay Development", level=1)

    for story_number, entry in enumerate(group["entries"], start=1):
        heading = document.add_heading(
            f"{story_number}. {_word_safe_text(entry['title']) or 'Untitled Story'}",
            level=2,
        )
        heading.paragraph_format.keep_with_next = True

        info = document.add_paragraph()
        info.paragraph_format.space_after = Pt(3)
        info_run = info.add_run(
            f"{entry['created_at'].astimezone(ZoneInfo(REPORT_TIMEZONE)).strftime('%Y-%m-%d')}"
            f" • {entry['url']}"
        )
        info_run.italic = True
        info_run.font.size = Pt(8)

        if not entry["messages"]:
            document.add_paragraph("No readable text or attachments were found.")
            continue

        for message_index, item in enumerate(entry["messages"]):
            if message_index > 0:
                continuation = document.add_paragraph()
                continuation.paragraph_format.space_before = Pt(4)
                continuation.paragraph_format.space_after = Pt(1)
                run = continuation.add_run(
                    "Author continuation — "
                    f"{item['created_at'].astimezone(ZoneInfo(REPORT_TIMEZONE)).strftime('%Y-%m-%d %H:%M')}"
                )
                run.bold = True
                run.font.size = Pt(8)

            if item["content"]:
                # Preserve Discord line breaks inside one Word paragraph. This
                # prevents every blank line from adding Word paragraph spacing.
                paragraph = document.add_paragraph()
                paragraph.paragraph_format.space_after = Pt(3)
                paragraph.add_run(item["content"])

            if item["attachments"]:
                attachment_paragraph = document.add_paragraph()
                attachment_paragraph.paragraph_format.space_after = Pt(2)
                attachment_paragraph.add_run("Attachments: ").bold = True
                attachment_paragraph.add_run(" | ".join(item["attachments"]))

    document.save(output_path)


def build_character_story_zip(groups, output_path):
    """Build every character DOCX and package them into one ZIP archive."""

    used_names = set()

    with tempfile.TemporaryDirectory(prefix="golden_age_story_report_") as temp_dir:
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for group in groups:
                if group["unresolved"]:
                    base_name = _safe_report_filename(
                        f"Unresolved_{group['author'].display_name}"
                    )
                else:
                    base_name = _safe_report_filename(
                        f"Character_{group['character_name']}_{group['race']}"
                    )
                filename = f"{base_name}.docx"
                suffix = 2

                while filename.casefold() in used_names:
                    filename = f"{base_name}_{suffix}.docx"
                    suffix += 1

                used_names.add(filename.casefold())
                document_path = os.path.join(temp_dir, filename)
                build_character_story_document(group, document_path)
                archive.write(document_path, arcname=filename)


async def generate_character_story_zip(guild: discord.Guild):
    """Generate a ZIP containing one compact DOCX per resolved character."""

    story_channel = await get_channel(CHARACTER_STORIES_CHANNEL_ID)
    if story_channel is None or story_channel.guild.id != guild.id:
        raise RuntimeError("The configured Character Stories channel could not be found.")

    entries, applications_by_author = await asyncio.gather(
        collect_story_entries(story_channel),
        collect_character_applications_for_story_report(),
    )
    groups = group_stories_by_character(entries, applications_by_author)

    os.makedirs(REPORTS_DIR, exist_ok=True)
    timestamp = datetime.now(ZoneInfo(REPORT_TIMEZONE)).strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(REPORTS_DIR, f"character_story_files_{timestamp}.zip")

    await asyncio.to_thread(build_character_story_zip, groups, output_path)

    unresolved_groups = sum(1 for group in groups if group["unresolved"])
    resolved_groups = len(groups) - unresolved_groups

    return {
        "path": output_path,
        "stories": len(entries),
        "documents": len(groups),
        "resolved": resolved_groups,
        "unresolved": unresolved_groups,
    }


async def generate_story_report(guild: discord.Guild):
    """Download all character stories and cross-reference applications."""

    story_channel = await get_channel(CHARACTER_STORIES_CHANNEL_ID)
    if story_channel is None or story_channel.guild.id != guild.id:
        raise RuntimeError("The configured Character Stories channel could not be found.")

    entries, applications_by_author = await asyncio.gather(
        collect_story_entries(story_channel),
        collect_character_applications_for_story_report(),
    )

    os.makedirs(REPORTS_DIR, exist_ok=True)
    timestamp = datetime.now(ZoneInfo(REPORT_TIMEZONE)).strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(REPORTS_DIR, f"character_stories_{timestamp}.docx")

    stats = get_story_report_stats(entries, applications_by_author)

    await asyncio.to_thread(
        build_story_report_document,
        entries,
        applications_by_author,
        output_path,
    )

    return {
        "path": output_path,
        "total": stats["stories"],
        "unique_players": stats["unique_players"],
        "accepted_characters": stats["accepted_characters"],
    }


async def handle_story_report_command(message: discord.Message):
    """Owner-only entry point for the per-character ZIP report."""

    if message.author.id != message.guild.owner_id:
        await message.reply("❌ Only the server owner can generate the story report.")
        return

    report_key = ("stories_zip", message.guild.id, message.author.id)
    if report_key in REPORT_IN_PROGRESS:
        await message.reply("⏳ A per-character story report is already being generated.")
        return

    REPORT_IN_PROGRESS.add(report_key)
    progress_message = None

    try:
        progress_message = await message.reply(
            "📚 Collecting Character Stories and building one Word file per character..."
        )
        result = await generate_character_story_zip(message.guild)
        await message.channel.send(
            "## 📚 Character Story Files Ready\n\n"
            f"**Stories collected:** `{result['stories']}`\n"
            f"**Word documents created:** `{result['documents']}`\n"
            f"**Characters matched automatically:** `{result['resolved']}`\n"
            f"**Players requiring manual review:** `{result['unresolved']}`\n\n"
            "The ZIP contains one compact Word document per character. "
            "Ambiguous cases are labeled `Unresolved` and were not assigned automatically.\n\n"
            "The ZIP was also saved in the bot's `reports/` folder.",
            file=discord.File(
                result["path"],
                filename=os.path.basename(result["path"]),
            ),
        )
        if progress_message is not None:
            try:
                await progress_message.delete()
            except discord.HTTPException:
                pass
    except Exception as error:
        print(f"❌ Per-character story report generation failed: {error!r}")
        error_text = (
            "❌ **Golden Age Manager could not generate the per-character story report.**\n\n"
            f"`{type(error).__name__}: {error}`"
        )
        if progress_message is not None:
            try:
                await progress_message.edit(content=error_text)
            except discord.HTTPException:
                await message.channel.send(error_text)
        else:
            await message.channel.send(error_text)
    finally:
        REPORT_IN_PROGRESS.discard(report_key)


async def handle_story_report_all_command(message: discord.Message):
    """Owner-only entry point for the legacy consolidated Word report."""

    if message.author.id != message.guild.owner_id:
        await message.reply("❌ Only the server owner can generate the story report.")
        return

    report_key = ("stories_all", message.guild.id, message.author.id)
    if report_key in REPORT_IN_PROGRESS:
        await message.reply("⏳ A consolidated Character Stories report is already being generated.")
        return

    REPORT_IN_PROGRESS.add(report_key)
    progress_message = None

    try:
        progress_message = await message.reply(
            "📖 Collecting Character Stories and building the consolidated Word document..."
        )
        result = await generate_story_report(message.guild)
        await message.channel.send(
            "## 📖 Consolidated Character Stories Report Ready\n\n"
            f"**Stories collected:** `{result['total']}`\n"
            f"**Unique players represented:** `{result['unique_players']}`\n"
            f"**Accepted characters represented:** `{result['accepted_characters']}`\n\n"
            "The Word file was also saved in the bot's `reports/` folder.",
            file=discord.File(
                result["path"],
                filename=os.path.basename(result["path"]),
            ),
        )
        if progress_message is not None:
            try:
                await progress_message.delete()
            except discord.HTTPException:
                pass
    except Exception as error:
        print(f"❌ Consolidated story report generation failed: {error!r}")
        error_text = (
            "❌ **Golden Age Manager could not generate the consolidated story report.**\n\n"
            f"`{type(error).__name__}: {error}`"
        )
        if progress_message is not None:
            try:
                await progress_message.edit(content=error_text)
            except discord.HTTPException:
                await message.channel.send(error_text)
        else:
            await message.channel.send(error_text)
    finally:
        REPORT_IN_PROGRESS.discard(report_key)

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

    if command in STORY_REPORT_COMMANDS:
        await handle_story_report_command(message)
        return

    if command in STORY_REPORT_ALL_COMMANDS:
        await handle_story_report_all_command(message)
        return

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
        == FORM_APPLICATION_FORUM_ID
    ):

        await handle_form_application(
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

    removed_tags = (
        before_tags - after_tags
    )

    approval_tag_names = {
        normalize(source_tag)
        for source_tag in MENTOR_APPROVAL_MAPPING
    }

    approval_removed = bool(
        removed_tags & approval_tag_names
    )

    # If Staff removes an approval tag, immediately remove that mentor from
    # the public directory. The approved profile itself is intentionally not
    # deleted automatically; restoring the approval tag will either reuse the
    # existing profile or rebuild it if Staff deleted it manually.
    if approval_removed:
        await rebuild_approved_mentor_records()
        await update_mentor_list_message(
            reason="mentor approval tag removed"
        )
        await update_available_abilities_message(
            reason="mentor approval tag removed"
        )

        if not newly_added:
            return

    if not newly_added:
        return


    # --------------------------------------------------------
    # REJECTED
    # --------------------------------------------------------

    if normalize(
        MENTOR_REJECTED_TAG
    ) in newly_added:

        closed = await require_decline_reason_and_reject(
            after,
            "Mentor Application",
            "mentor",
            MENTOR_REJECTED_TAG,
            normalizer=normalize,
        )

        if closed:
            await rebuild_approved_mentor_records()
            await update_mentor_list_message(
                reason="mentor application rejected"
            )
            await update_available_abilities_message(
                reason="mentor application rejected"
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
    # PAIRED TECHNIQUE REQUIREMENT
    # --------------------------------------------------------
    # Some techniques are bundled into one mentor slot. If Staff attempts to
    # approve only one technique from a required pair, remove the approval tag
    # and leave the application open so it can be corrected.
    validation_starter = await get_starter_message(after)
    if validation_starter is None:
        print("❌ Starter message not found for mentor bundle validation.")
        return

    validation_application = parse_mentor_application(
        validation_starter.content or ""
    )
    validation_skill = (
        validation_application.get("skill")
        or after.name
        or ""
    )

    bundle_error = _mentor_bundle_requirement_error(
        validation_skill,
        destination_tag_name,
    )

    if bundle_error:
        try:
            selected_normalized = normalize(selected_source_tag)
            remaining_tags = [
                tag
                for tag in after.applied_tags
                if normalize(tag.name) != selected_normalized
            ]

            await after.edit(
                applied_tags=remaining_tags,
                reason=(
                    "Golden Age Manager: incomplete paired mentor bundle"
                ),
            )

            await after.send(
                "## ⚠️ Mentor Bundle Requirement\n\n"
                f"{bundle_error}\n\n"
                "The approval tag was removed automatically. "
                "The application was **not published or archived**."
            )

            print(
                f"⚠️ Mentor approval blocked for '{after.name}': "
                "paired technique bundle requirement not met."
            )

        except discord.Forbidden:
            print(
                "❌ Could not remove the mentor approval tag after the "
                "paired mentor bundle validation failed. Check Manage Threads."
            )
        except discord.HTTPException as error:
            print(
                f"❌ Discord API error while enforcing paired mentor bundle: {error}"
            )

        await rebuild_approved_mentor_records()
        await update_mentor_list_message(
            reason="incomplete paired mentor bundle blocked"
        )
        await update_available_abilities_message(
            reason="incomplete paired mentor bundle blocked"
        )
        return


    # --------------------------------------------------------
    # DUPLICATE PROTECTION + SELF-HEALING
    # --------------------------------------------------------

    if application_was_processed(
        "mentor",
        after.id
    ):

        destination_exists = await processed_destination_exists(
            "mentor",
            after.id,
            APPROVED_MENTORS_FORUM_ID
        )

        if destination_exists:
            await rebuild_approved_mentor_records()
            await update_mentor_list_message(
                reason="mentor approval tag restored"
            )
            await update_available_abilities_message(
                reason="mentor approval tag restored"
            )

            print(
                "⚠️ This mentor application was already processed. "
                "Profile duplicate skipped; Mentor List refreshed."
            )

            return

        # The source still has a stale processed marker but the copied
        # approved-mentors profile was deleted. Clear the marker and continue
        # through the normal approval flow so the profile is recreated.
        unmark_application_processed(
            "mentor",
            after.id
        )

        print(
            "♻️ Processed mentor profile is missing from approved-mentors. "
            "Rebuilding it now..."
        )


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

        await rebuild_approved_mentor_records()
        await update_mentor_list_message(
            reason="mentor approved"
        )
        await update_available_abilities_message(
            reason="mentor approved"
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

        closed = await require_decline_reason_and_reject(
            after,
            "Character Application",
            "character",
            CHARACTER_REJECTED_TAG,
            normalizer=normalize,
        )

        if closed:
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
                "`Race: Tuffle`\n"
                "`Race: Grey Race`\n"
                "`Race: Demon Race`\n"
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
# FORM APPLICATION AUTOMATION
# ============================================================

async def handle_form_application(
    before: discord.Thread,
    after: discord.Thread,
):
    """Handle Accepted / Rejected / Voided states in #form-application."""

    before_tags = {
        normalize_forum_tag(tag.name)
        for tag in before.applied_tags
    }
    after_tags = {
        normalize_forum_tag(tag.name)
        for tag in after.applied_tags
    }

    newly_added = after_tags - before_tags
    if not newly_added:
        return

    accepted_name = normalize_forum_tag(FORM_ACCEPTED_TAG)
    rejected_name = normalize_forum_tag(FORM_REJECTED_TAG)
    voided_name = normalize_forum_tag(FORM_VOIDED_TAG)

    # --------------------------------------------------------
    # REJECTED
    # --------------------------------------------------------
    if rejected_name in newly_added:
        await require_decline_reason_and_reject(
            after,
            "Form Application",
            "form",
            FORM_REJECTED_TAG,
            status_word="Rejected",
            normalizer=normalize_forum_tag,
        )
        return

    # --------------------------------------------------------
    # VOIDED
    # --------------------------------------------------------
    if voided_name in newly_added:
        await require_decline_reason_and_reject(
            after,
            "Form Application",
            "form",
            FORM_VOIDED_TAG,
            status_word="Voided",
            normalizer=normalize_forum_tag,
        )
        return

    # --------------------------------------------------------
    # ACCEPTED
    # --------------------------------------------------------
    if accepted_name not in newly_added:
        return

    if application_was_processed("form", after.id):
        destination_exists = await processed_destination_exists(
            "form",
            after.id,
            APPROVED_FORMS_FORUM_ID,
        )

        if destination_exists:
            print(
                "⚠️ This Form Application was already processed. "
                "Approved-form duplicate skipped."
            )
            return

        unmark_application_processed("form", after.id)
        print(
            "♻️ Processed Form profile is missing from approved-forms. "
            "Rebuilding it now..."
        )

    print()
    print("=" * 65)
    print("🔥 FORM APPROVAL DETECTED")
    print(f"📋 {after.name}")
    print("=" * 65)

    approved_forum = await get_channel(APPROVED_FORMS_FORUM_ID)
    if not isinstance(approved_forum, discord.ForumChannel):
        print("❌ approved-forms not found.")
        return

    starter_message = await get_starter_message(after)
    if starter_message is None:
        print("❌ Starter message not found for Form Application.")
        return

    application = parse_form_application(starter_message.content or "")
    character_name = clean_value(application.get("character_name") or "")
    current_level = clean_value(application.get("current_level") or "")
    time_played = clean_value(application.get("time_played") or "")
    form_requested = clean_value(application.get("form_requested") or "")
    reasoning = clean_value(application.get("reasoning") or "")
    roleplay_plan = clean_value(application.get("roleplay_plan") or "")

    embed = discord.Embed(title="✅ Approved Form")
    embed.set_author(
        name=starter_message.author.display_name,
        icon_url=starter_message.author.display_avatar.url,
    )

    if character_name and form_requested:
        embed.description = f"## {character_name}\n**{form_requested}**"
    elif character_name:
        embed.description = f"## {character_name}"
    elif form_requested:
        embed.description = f"## {form_requested}"

    if current_level:
        embed.add_field(
            name="⚡ Current Level",
            value=current_level[:1024],
            inline=True,
        )

    if time_played:
        embed.add_field(
            name="⏱️ Time Played",
            value=time_played[:1024],
            inline=True,
        )

    if form_requested:
        embed.add_field(
            name="🔥 Form",
            value=form_requested[:1024],
            inline=False,
        )

    if reasoning:
        add_long_field(embed, "📖 Reasoning", reasoning)

    if roleplay_plan:
        add_long_field(embed, "🎭 Roleplay Plan", roleplay_plan)

    embed.add_field(
        name="🔗 Application",
        value=f"[View original application]({starter_message.jump_url})",
        inline=False,
    )

    embed.set_footer(
        text="Dragon Ball: Golden Age • Form Unlocking Program"
    )

    files, thumbnail_filename = await prepare_attachments(starter_message)
    if thumbnail_filename:
        embed.set_thumbnail(
            url=f"attachment://{thumbnail_filename}"
        )

    try:
        # approved-forms currently does not require a category mapping from the
        # source application. If Staff later adds destination tags (Saiyan,
        # Namekian, etc.), a mapping can be added without changing this flow.
        result = await approved_forum.create_thread(
            name=after.name[:100],
            embed=embed,
            files=files,
        )

        approved_thread = result.thread
        mark_application_processed(
            "form",
            after.id,
            approved_thread.id,
        )

        await after.send(
            "## ✅ Form Application Approved\n\n"
            "Your Form Application has been accepted by the Staff Team.\n\n"
            f"🔥 **Form:** `{form_requested or 'See application'}`\n\n"
            f"Your approved record has been published in "
            f"{approved_forum.mention}."
        )

        await after.edit(
            archived=True,
            locked=True,
            reason="Golden Age Manager: Form Application approved",
        )

        print(f"✅ Approved form profile created: {approved_thread.name}")
        print("🔒 Original Form Application archived.")

    except discord.Forbidden:
        print(
            "❌ Permission error during Form approval. Check Create Posts, "
            "Send Messages in Posts, Embed Links, Attach Files, Manage Posts, "
            "and Read Message History."
        )

    except discord.HTTPException as error:
        print(f"❌ Discord API error during Form approval: {error}")

    except Exception as error:
        print(f"❌ Unexpected Form approval error: {error!r}")


# ============================================================
# MUTANT AUTOMATION
# ============================================================

async def handle_mutant_application(
    before: discord.Thread,
    after: discord.Thread,
):
    """Processes the dedicated Mutant Application forum."""

    before_tags = {normalize_forum_tag(tag.name) for tag in before.applied_tags}
    after_tags = {normalize_forum_tag(tag.name) for tag in after.applied_tags}
    newly_added = after_tags - before_tags

    accepted_name = normalize_forum_tag(MUTANT_ACCEPTED_TAG)
    rejected_name = normalize_forum_tag(MUTANT_REJECTED_TAG)

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
                    if normalize_forum_tag(tag.name) != accepted_name
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

    normalized_after_tags = {normalize_forum_tag(tag.name) for tag in after.applied_tags}

    if normalize_forum_tag(MUTANT_MIGRATION_TAG) in normalized_after_tags:
        application_type = "Migration"
    elif normalize_forum_tag(MUTANT_NEW_APPLICATION_TAG) in normalized_after_tags:
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