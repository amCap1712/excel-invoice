from datetime import datetime, date
from enum import Enum

from Levenshtein import distance
from pandas import DataFrame, Timestamp


class Restaurant(Enum):
    DAWAT = "Dawat"
    WELCOME_INDIA = "WelcomeIndia"


RESTAURANTS = [Restaurant.DAWAT, Restaurant.WELCOME_INDIA]


CANONICAL_DMCS = {
    "Neemholidays": "Neem Holidays",
    "Star Our": "Star Tour",
    "Star": "Star Tour",
    "Ezxa": "Exza",
    "Tc": "TC",
    "Gtt": "GTT",
    "Chr": "CHR",
    "Gb Dmc": "GB DMC",
    "Gb Dmc Ltd.": "GB DMC",
    "Youngedsplorer": "Young Edsplorer",
    "Truvaiglobal": "Truvai",
    "Truvai Dmc": "Truvai",
    "Truvai Global": "Truvai",
    "Europe Incomign": "Europe Incoming",
    "Europeincoming": "Europe Incoming",
    "Afc Holidyays": "AFC Holidays",
    "Afc Holidays": "AFC Holidays",
    "G2 Travel": "G2 Travels",
    "G2Travel": "G2 Travels",
    "G2-Travel": "G2 Travels",
    "Holiday Carnival": "Holidays Carnival",
    "Europe Goodlife": "Europe Good Life",
    "Europegoodlife": "Europe Good Life",
    "Gateways Group Of Dmcs": "Gateways Group Of Dmc'S",
    "Gateways Dmc": "Gateways Group Of Dmc'S",
    "Gateways Group Of Dmc’S": "Gateways Group Of Dmc'S",
    "Deewan Holidays": "Dewan Holidays",
    "Dewan Travels": "Dewan Holidays",
    "Switru": "Switrus",
    "European Gatewyas": "European Gateways",
    "Lamour Voyage": "Lamour Voyages",
    "Lamondialetour": "La Mondiale",
    "Lamondiale Tour": "La Mondiale",
    "Lamondiale Tours": "La Mondiale",
    "Mahadevan Group": "Mahadevan",
    "Whats App": "WhatsApp",
    "Whatassp": "WhatsApp",
    "Whatsapp": "WhatsApp"
}
IGNORE_DMC_DUPES = [("TC", "Tcf"), ("Magi Holidays", "Mango Holidays"), ("GTT", "TC")]


def convert_to_date(value):
    if isinstance(value, Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def find_possible_dupes(updater, dmcs):
    for dmc1 in dmcs:
        dupes = []
        for dmc2 in dmcs:
            if dmc1 == dmc2:
                continue
            if distance(dmc1, dmc2) <= 2\
                    and (dmc1, dmc2) not in IGNORE_DMC_DUPES\
                    and (dmc2, dmc1) not in IGNORE_DMC_DUPES:
                dupes.append(dmc2)
        if dupes:
            updater(f"Possible duplicates for '{dmc1}': {dupes}")


def process_data(updater, from_date, to_date, df: DataFrame) -> tuple[DataFrame, DataFrame]:
    df["Service Date Cleaned"] = df["Service Date"].apply(convert_to_date)

    df = df.loc[(df["Service Date Cleaned"] >= from_date) & (df["Service Date Cleaned"] <= to_date)]

    df["Adult"] = df["Adult"].fillna(0)
    df["Children"] = df["Children"].fillna(0)

    df["Dmc Canonical"] = df["Dmc"]
    df["Dmc Canonical"] = df["Dmc Canonical"].str.title().str.strip().str.split().str.join(" ")
    df["Dmc Canonical"] = df["Dmc Canonical"].apply(lambda dmc: CANONICAL_DMCS.get(dmc, dmc))

    # TODO: Create a list of valid DMCs, write invalid ones to a separate file
    unique_dmcs = df["Dmc Canonical"].unique().tolist()
    if None in unique_dmcs:
        unique_dmcs.remove(None)
    find_possible_dupes(updater, unique_dmcs)

    # TODO: read dmc wide and global prices from excel sheet
    if "Price Adult" in df.columns:
        df["Price Adult"] = df["Price Adult"].fillna(13)

    if "Price Child" in df.columns:
        df["Price Child"] = df["Price Child"].fillna(8)

    if "Remarks" in df.columns:
        cancelled_mask_1 = df["Remarks"].fillna("").str.lower().str.startswith("cancel")
    else:
        cancelled_mask_1 = True

    if "Delivery" in df.columns:
        cancelled_mask_2 = df["Delivery"].fillna("").str.lower().str.startswith("cancel")
    else:
        cancelled_mask_2 = True

    cancelled_mask = cancelled_mask_1 | cancelled_mask_2

    serviced_df = df[~cancelled_mask]
    cancelled_df = df[cancelled_mask]

    return serviced_df, cancelled_df
