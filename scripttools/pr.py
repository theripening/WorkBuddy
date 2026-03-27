import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path


def run_for_today(card_file, ach_file, ta_file):
    """
    Accepts three Django UploadedFile objects, runs the PR merge for today's
    cutoff window, and returns (csv_bytes, filename).
    """
    import pandas as pd

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        for uploaded, name in [(card_file, "Card.xlsx"), (ach_file, "ACH.xlsx"), (ta_file, "TA.xlsx")]:
            dest = tmpdir / name
            with open(dest, "wb") as f:
                for chunk in uploaded.chunks():
                    f.write(chunk)

        today = date.today()
        start_dt, end_dt = _cutoff_window(today, "07:00")

        card = _merge(
            pd, tmpdir / "Card.xlsx", tmpdir / "TA.xlsx",
            join_col_a="InfoSend Transaction ID", join_col_b="Confirmation #",
            name_col="Cardholder Name", method_col="Card Type",
            start_date=start_dt, end_date=end_dt,
        )
        ach = _merge(
            pd, tmpdir / "ACH.xlsx", tmpdir / "TA.xlsx",
            join_col_a="InfoSend Transaction ID", join_col_b="Confirmation #",
            name_col="Account Holder Name", method_col="Payment Type",
            start_date=start_dt, end_date=end_dt,
        )

        final = pd.concat([card, ach], ignore_index=True)
        final = final.sort_values(by="InfoSend Transaction ID")
        csv_bytes = final.to_csv(index=False, header=False).encode("utf-8")

    filename = f"{today.strftime('%Y%m%d')}_output.csv"
    return csv_bytes, filename


def _cutoff_window(target_date, cutoff_time="07:00"):
    h, m = map(int, cutoff_time.split(":"))
    end_dt = datetime.combine(target_date, datetime.min.time()).replace(hour=h, minute=m)
    return end_dt - timedelta(days=1), end_dt


def _merge(pd, file_a, file_b, join_col_a, join_col_b, name_col, method_col,
           start_date=None, end_date=None):
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    df_a.columns = df_a.columns.str.strip()
    df_b.columns = df_b.columns.str.strip()

    df_a["Transaction Date"] = pd.to_datetime(
        df_a["Transaction Date"], format="%m/%d/%y  %I:%M:%S %p", errors="coerce"
    )
    if start_date:
        df_a = df_a[df_a["Transaction Date"] >= start_date]
    if end_date:
        df_a = df_a[df_a["Transaction Date"] < end_date]

    merged = pd.merge(df_a, df_b, left_on=join_col_a, right_on=join_col_b, how="inner")

    acct_col = "Account Number (XXXXXXX-XXXXXX)"
    merged["Account First Part"] = merged[acct_col].astype(str).str.split("-").str[0]
    merged["Account Second Part"] = merged[acct_col].astype(str).str.split("-").str[1]
    merged["First Name"] = merged[name_col].astype(str).str.split().str[0]
    merged["Last Name"] = merged[name_col].astype(str).str.split().str[-1]
    merged["Method"] = merged[method_col] if method_col in merged.columns else None

    return merged[[
        join_col_a, "Account First Part", "Account Second Part",
        "Amount", "Method", "First Name", "Last Name",
    ]]
