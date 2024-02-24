from datetime import datetime, date
from enum import Enum

from Levenshtein import distance
from pandas import DataFrame, Timestamp, concat


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


def process_rates_df(updater, rates_df: DataFrame):
    rates_df = rates_df.rename(columns={"DMC": "Dmc Canonical"})
    rates_df["Dmc Canonical"] = rates_df["Dmc Canonical"].str.strip()

    default_rates = rates_df.loc[rates_df["Dmc Canonical"] == "Default"]
    if default_rates.empty:
        updater("Default rates not found")
        return rates_df
    default_rates = default_rates.to_dict(orient="records")[0]
    default_rates.pop("Dmc Canonical")
    rates_df = rates_df.fillna(default_rates)

    rates_df = rates_df.melt(id_vars=["Dmc Canonical"], var_name="Service Type", value_name="Rate")
    child_rates_df = rates_df.loc[rates_df["Service Type"] == "Child"]
    child_rates_df = child_rates_df.drop(columns=["Service Type"])
    rates_df = rates_df.merge(child_rates_df, on="Dmc Canonical", how="outer", suffixes=(None, " Child"))

    rates_df["Dmc To Join"] = rates_df["Dmc Canonical"].str.lower().str.strip().str.split().str.join(" ")

    return rates_df


def filter_cancelled_tours(df: DataFrame) -> tuple[DataFrame, DataFrame]:
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


def filter_unknown_dmcs(df: DataFrame, rates_df: DataFrame) -> tuple[DataFrame, DataFrame]:
    unique_dmcs = set(rates_df["Dmc To Join"].unique().tolist())
    df["Dmc To Join"] = df["Dmc"].str.lower().str.strip().str.split().str.join(" ")

    known_dmcs_mask = df["Dmc To Join"].isin(unique_dmcs)
    known_dmcs_df = df[known_dmcs_mask]
    unknown_dmcs_df = df[~known_dmcs_mask]
    return known_dmcs_df, unknown_dmcs_df


def filter_unknown_rates(df: DataFrame, rates_df: DataFrame) -> tuple[DataFrame, DataFrame]:
    df["Service Type"] = df["Service Type"].str.title().str.strip().str.split().str.join(" ")

    df["Dmc To Join"] = df["Dmc"].str.lower().str.strip().str.split().str.join(" ")
    df = df.merge(rates_df, how="left", on=["Dmc To Join", "Service Type"])

    unknown_mask = (df["Rate"].isna() | df["Rate"].isnull()) & df["Price Adult"].isna() & df["Price Child"].isna()
    unknown_df = df.loc[unknown_mask]
    known_df = df.loc[~unknown_mask]

    known_df["Price Adult"] = known_df["Price Adult"].fillna(known_df["Rate"])
    known_df["Price Child"] = known_df["Price Child"].fillna(known_df["Rate Child"])

    return known_df, unknown_df


def filter_unknown_dates(df: DataFrame) -> tuple[DataFrame, DataFrame]:
    df["Service Date Cleaned"] = df["Service Date"].apply(convert_to_date)
    invalid_date_mask = df["Service Date Cleaned"].isnull()
    invalid_dates_df = df[invalid_date_mask]
    valid_dates_df = df[~invalid_date_mask]
    return valid_dates_df, invalid_dates_df


def filter_missing_counts(df: DataFrame) -> tuple[DataFrame, DataFrame]:
    missing_counts_mask = df["Adult"].isna() & df["Children"].isna()
    missing_counts_df = df[missing_counts_mask]
    valid_counts_df = df[~missing_counts_mask]
    valid_counts_df["Adult"] = valid_counts_df["Adult"].fillna(0)
    valid_counts_df["Children"] = valid_counts_df["Children"].fillna(0)
    return valid_counts_df, missing_counts_df


def fixup_invalid_df(invalid_df: DataFrame, reason: str) -> DataFrame | None:
    invalid_df = invalid_df.dropna(axis=1, how="all")
    if invalid_df.empty:
        return None
    invalid_df["Reason"] = reason
    invalid_df.insert(0, "Reason", invalid_df.pop("Reason"))
    return invalid_df


def process(from_date, to_date, rates_df: DataFrame, df: DataFrame) -> tuple[DataFrame, DataFrame, DataFrame]:
    df, unknown_dates_df = filter_unknown_dates(df)
    df = df.loc[(df["Service Date Cleaned"] >= from_date) & (df["Service Date Cleaned"] <= to_date)]

    df, cancelled_df = filter_cancelled_tours(df)
    df, unknown_dmcs_df = filter_unknown_dmcs(df, rates_df)
    df, unknown_rates_df = filter_unknown_rates(df, rates_df)
    df, missing_counts_df = filter_missing_counts(df)

    unknown_dates_df = fixup_invalid_df(unknown_dates_df, "Service date could not be parsed")
    unknown_dmcs_df = fixup_invalid_df(unknown_dmcs_df, "DMC is not known")
    unknown_rates_df = fixup_invalid_df(unknown_rates_df, "Service type is unknown and Price Adult/Child not defined")
    missing_counts_df = fixup_invalid_df(missing_counts_df, "Both adult and children count is missing")

    invalid_dfs = []
    if unknown_dates_df is not None:
        invalid_dfs.append(unknown_dates_df)
    if unknown_dmcs_df is not None:
        invalid_dfs.append(unknown_dmcs_df)
    if unknown_rates_df is not None:
        invalid_dfs.append(unknown_rates_df)
    if missing_counts_df is not None:
        invalid_dfs.append(missing_counts_df)

    if invalid_dfs:
        invalid_df = concat(invalid_dfs, ignore_index=True)
    else:
        invalid_df = DataFrame()

    return df, cancelled_df, invalid_df
