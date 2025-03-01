from dataclasses import dataclass
from datetime import datetime, date
from textwrap import dedent

import numpy as np
import pandas
from dateutil import parser
from pandas import DataFrame, Timestamp, concat

pandas.set_option("display.max_rows", None)
pandas.set_option("display.max_columns", None)
pandas.set_option("display.width", None)


@dataclass
class Restaurant:
    name: str
    address: str


RESTAURANTS = [
    Restaurant("Dawat", dedent("""
        Dawat
        58 Avenue du 8 mai 1945
        93150 Le Blanc Mesnil Paris, France.
    """).strip()),
    Restaurant("WelcomeIndia", dedent("""
        Welcome India
        9 Bis Boulevard Du Montparnasse
        75006 Paris, France.
    """).strip()),
    Restaurant("WaytoIndia", dedent("""
        Way to India
        Strombeeklinde 92, 1853 Grimbergen
        Bruxelles, Belgique.
    """).strip()),
    Restaurant("Tara", dedent("""
        Tara
        58 Avenue du 8 mai 1945
        93150 Le Blanc Mesnil Paris, France.
    """))
]


def convert_to_date(value):
    if isinstance(value, Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip().replace("_", "-")
        try:
            return parser.parse(value).date()
        except ValueError:
            pass
    return None


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
        cancelled_mask_1 = False

    if "Delivery" in df.columns:
        cancelled_mask_2 = df["Delivery"].fillna("").str.lower().str.startswith("cancel")
    else:
        cancelled_mask_2 = False

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


def filter_unknown_service_types(df: DataFrame, rates_df: DataFrame) -> tuple[DataFrame, DataFrame]:
    unique_service_types = set(rates_df["Service Type"].unique().tolist())
    df["Service Type"] = df["Service Type"].str.title().str.strip().str.split().str.join(" ")
    known_service_types_mask = df["Service Type"].isin(unique_service_types)
    known_service_types_df = df[known_service_types_mask]
    unknown_service_types_df = df[~known_service_types_mask]
    return known_service_types_df, unknown_service_types_df


def filter_unknown_rates(df: DataFrame, rates_df: DataFrame) -> tuple[DataFrame, DataFrame]:
    df["Dmc To Join"] = df["Dmc"].str.lower().str.strip().str.split().str.join(" ")
    df = df.merge(rates_df, how="left", on=["Dmc To Join", "Service Type"])

    if "Price Child" not in df.columns:
        df["Price Child"] = 0
    if "Price Adult" not in df.columns:
        df["Price Adult"] = 0

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
    missing_counts_mask = np.zeros(df.shape[0], dtype=bool)  # disable missing counts check for now
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


def process(restaurant: Restaurant, from_date, to_date, rates_df: DataFrame, df: DataFrame, not_found_df: DataFrame,
            typos_df: DataFrame) -> tuple[DataFrame, DataFrame, DataFrame]:
    df = df[df["Restaurant"] == restaurant.name]
    if "Tour Code" in df.columns:
        df["Tour Code"] = df["Tour Code"].astype("str").str.strip()

    df, unknown_dates_df = filter_unknown_dates(df)
    df = df.loc[(df["Service Date Cleaned"] >= from_date) & (df["Service Date Cleaned"] <= to_date)]

    df, cancelled_df = filter_cancelled_tours(df)
    df, unknown_dmcs_df = filter_unknown_dmcs(df, rates_df)
    df, unknown_service_types_df = filter_unknown_service_types(df, rates_df)
    df, unknown_rates_df = filter_unknown_rates(df, rates_df)
    df, missing_counts_df = filter_missing_counts(df)

    restaurant_not_found_df = not_found_df.loc[not_found_df["Restaurant"] == restaurant.name]
    restaurant_not_found_df = fixup_invalid_df(restaurant_not_found_df, "Sheet for restaurant not found")

    restaurant_typos_df = typos_df.loc[typos_df["Restaurant"] == restaurant.name]
    restaurant_typos_df = fixup_invalid_df(restaurant_typos_df, "Sheetname for restaurant has a typo")

    unknown_dates_df = fixup_invalid_df(unknown_dates_df, "Service date could not be parsed")
    unknown_dmcs_df = fixup_invalid_df(unknown_dmcs_df, "DMC is not known")
    unknown_service_types_df = fixup_invalid_df(unknown_service_types_df, "Service type is not known")
    unknown_rates_df = fixup_invalid_df(unknown_rates_df, "Service type is unknown and Price Adult/Child not defined")
    missing_counts_df = fixup_invalid_df(missing_counts_df, "Both adult and children count is missing")

    invalid_dfs = []
    if restaurant_not_found_df is not None and not restaurant_not_found_df.empty:
        invalid_dfs.append(restaurant_not_found_df)
    if restaurant_typos_df is not None and not restaurant_typos_df.empty:
        invalid_dfs.append(restaurant_typos_df)
    if unknown_dates_df is not None and not unknown_dates_df.empty:
        invalid_dfs.append(unknown_dates_df)
    if unknown_dmcs_df is not None and not unknown_dmcs_df.empty:
        invalid_dfs.append(unknown_dmcs_df)
    if unknown_service_types_df is not None and not unknown_service_types_df.empty:
        invalid_dfs.append(unknown_service_types_df)
    if unknown_rates_df is not None and unknown_rates_df.empty:
        invalid_dfs.append(unknown_rates_df)
    if missing_counts_df is not None and not missing_counts_df.empty:
        invalid_dfs.append(missing_counts_df)

    if invalid_dfs:
        invalid_df = concat(invalid_dfs, ignore_index=True)
    else:
        invalid_df = DataFrame()

    return df, cancelled_df, invalid_df
