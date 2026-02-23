#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from pathlib import Path
import math

# ===== Adjust these to match your sheet headers =====
COLS = {
    "id":   "task_code",
    "name": "task_name",
    "dur":  "target_drtn_hr_cnt",   # original duration (working days)
    "lp":   "driving_path_flag"     # Y = on Longest Path
}

CRASHED_COL = "Crashed Duration"
ROUND_TO_DAYS = True   # set False if you want decimals

def to_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default

def main():
    # 1) Pick file
    tk.Tk().withdraw()
    in_path = filedialog.askopenfilename(
        title="Select Primavera Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not in_path:
        return

    df = pd.read_excel(in_path)

    # 2) Validate required columns
    for key in ["dur", "lp"]:
        if COLS[key] not in df.columns:
            messagebox.showerror("Error", f"Missing column: {COLS[key]}")
            return

    # 3) Longest Path mask (Y/YES/1/TRUE)
    lp_mask = df[COLS["lp"]].astype(str).str.upper().isin(["Y","YES","1","TRUE"])

    # 4) Compute current project duration as sum of LP durations
    cur = float(df.loc[lp_mask, COLS["dur"]].apply(to_float).sum())

    # If nothing on LP, stop
    if cur <= 0:
        messagebox.showerror("Error", "No Longest Path durations found (sum = 0).")
        return

    # 5) Ask for target
    target = simpledialog.askfloat(
        "Target Project Duration",
        f"Current project duration (Longest Path) ≈ {cur:.1f} working days.\n\n"
        f"Enter target project duration (days):",
        minvalue=1.0,
        initialvalue=max(1.0, round(cur * 0.8, 1))
    )
    if target is None:
        return

    # 6) Compute proportional factor
    # New LP sum should equal 'target' → scale LP durations by ratio = target / current
    ratio = target / cur
    # Safety guard
    ratio = max(0.0, ratio)

    # 7) Build 'Crashed Duration' column (copy original first)
    df[CRASHED_COL] = df[COLS["dur"]].apply(to_float)

    # Scale only LP rows
    scaled = df.loc[lp_mask, COLS["dur"]].apply(to_float) * ratio
    if ROUND_TO_DAYS:
        scaled = scaled.round()  # round to nearest day

    df.loc[lp_mask, CRASHED_COL] = scaled

    # 8) Put the new column right after the duration column
    cols = list(df.columns)
    if CRASHED_COL in cols:
        cols.remove(CRASHED_COL)
    insert_at = cols.index(COLS["dur"]) + 1
    cols.insert(insert_at, CRASHED_COL)
    df = df[cols]

    # 9) Recompute achieved (LP sum after scaling)
    achieved = float(df.loc[lp_mask, CRASHED_COL].apply(to_float).sum())

    # 10) Save
    save_path = filedialog.asksaveasfilename(
        title="Save Crashed File",
        defaultextension=".xlsx",
        initialfile=Path(in_path).stem + "_lp_proportional_crashed.xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not save_path:
        return

    with pd.ExcelWriter(save_path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Crashed")
        summary = pd.DataFrame({
            "Metric": [
                "CurrentDuration(LongestPath)",
                "TargetDuration",
                "AchievedDuration(LongestPath)",
                "ScalingRatio (Target/Current)"
            ],
            "Value": [
                round(cur,2),
                round(target,2),
                round(achieved,2),
                round(ratio,4)
            ]
        })
        summary.to_excel(w, index=False, sheet_name="Summary")

    messagebox.showinfo(
        "Done",
        f"Current ≈ {cur:.0f}d → Target {target:.0f}d\n"
        f"Achieved ≈ {achieved:.0f}d\n\n"
        f"Saved:\n{save_path}"
    )

if __name__ == "__main__":
    main()
