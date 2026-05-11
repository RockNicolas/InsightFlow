from dataclasses import dataclass

from modules.fleet_processors import itarema, saneamento


@dataclass(frozen=True)
class FleetProfile:
    key: str
    label: str
    upload_subfolder: str
    output_subfolder: str
    processor: object


FLEET_PROFILES = {
    "saneamento": FleetProfile(
        key="saneamento",
        label="Saneamento",
        upload_subfolder="saneamento",
        output_subfolder="saneamento",
        processor=saneamento,
    ),
    "itarema": FleetProfile(
        key="itarema",
        label="Itarema",
        upload_subfolder="itarema",
        output_subfolder="itarema",
        processor=itarema,
    ),
}


DEFAULT_FLEET_KEY = "saneamento"


def list_fleet_profiles():
    return list(FLEET_PROFILES.values())


def get_fleet_profile(profile_key):
    return FLEET_PROFILES.get(profile_key, FLEET_PROFILES[DEFAULT_FLEET_KEY])
