import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import warnings
import os

# import missingno as msno


rumus_grammar = pd.DataFrame(
    {
        "g1": [
            0,
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            9,
            10,
            11,
            12,
            13,
            14,
            15,
            16,
            17,
            18,
            19,
            20,
            21,
            22,
            23,
            24,
            25,
            26,
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            35,
            36,
            37,
            38,
            39,
            40,
        ],
        "g2": [
            21,
            21,
            22,
            24,
            25,
            26,
            28,
            29,
            30,
            32,
            33,
            34,
            35,
            36,
            37,
            38,
            39,
            40,
            41,
            42,
            43,
            44,
            45,
            46,
            47,
            48,
            49,
            50,
            51,
            52,
            53,
            54,
            55,
            57,
            58,
            59,
            61,
            63,
            64,
            66,
            67,
        ],
    }
)

rumus_listening = pd.DataFrame(
    {
        "l1": [
            0,
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            9,
            10,
            11,
            12,
            13,
            14,
            15,
            16,
            17,
            18,
            19,
            20,
            21,
            22,
            23,
            24,
            25,
            26,
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            35,
            36,
            37,
            38,
            39,
            40,
            41,
            42,
            43,
            44,
            45,
            46,
            47,
            48,
            49,
            50,
        ],
        "l2": [
            25,
            25,
            26,
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            36,
            37,
            38,
            39,
            40,
            41,
            41,
            42,
            43,
            43,
            44,
            44,
            45,
            46,
            46,
            47,
            48,
            48,
            49,
            49,
            50,
            50,
            51,
            52,
            52,
            53,
            54,
            55,
            56,
            56,
            57,
            58,
            59,
            60,
            61,
            62,
            63,
            64,
            66,
            68,
        ],
    }
)

rumus_reading = pd.DataFrame(
    {
        "r1": [
            0,
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            9,
            10,
            11,
            12,
            13,
            14,
            15,
            16,
            17,
            18,
            19,
            20,
            21,
            22,
            23,
            24,
            25,
            26,
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            35,
            36,
            37,
            38,
            39,
            40,
            41,
            42,
            43,
            44,
            45,
            46,
            47,
            48,
            49,
            50,
        ],
        "r2": [
            22,
            22,
            23,
            24,
            25,
            26,
            27,
            28,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            35,
            36,
            37,
            38,
            39,
            40,
            41,
            41,
            42,
            43,
            44,
            45,
            45,
            46,
            47,
            47,
            48,
            48,
            49,
            49,
            50,
            51,
            52,
            53,
            54,
            55,
            56,
            57,
            58,
            59,
            60,
            61,
            63,
            65,
            66,
            67,
        ],
    }
)


def select_files():
    progress_bar.start()
    filenames = filedialog.askopenfilenames(
        title="Select Excel Files", filetypes=[("Excel Files", "*.xlsx")]
    )
    file_paths = []
    progress_bar["value"] = 10
    for filename in filenames:
        file_paths.append(filename)
    for widget in root.pack_slaves():
        if isinstance(widget, tk.Label):
            widget.destroy()
    progress_bar["value"] = 20
    create_table(file_paths)
    progress_bar["value"] = 30
    return file_paths


def create_table(file_paths):
    progress_bar["value"] = 40
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        df = pd.concat([pd.read_excel(file_path) for file_path in file_paths])
    progress_bar["value"] = 50
    df = df[["No. Peserta", "Section 1", "Section 2", "Section 3"]]
    df.columns = ["ID", "l1", "g1", "r1"]
    progress_bar["value"] = 60
    df = df[(df["l1"] != 0) | (df["g1"] != 0) | (df["r1"] != 0)]
    # print(df.info())
    df = df.merge(rumus_listening, on="l1", how="left")
    df = df.merge(rumus_grammar, on="g1", how="left")
    df = df.merge(rumus_reading, on="r1", how="left")
    progress_bar["value"] = 70
    df["skor"] = round((df["l2"] + df["g2"] + df["r2"]) * 10 / 3)
    df.loc[df["skor"] < 310, "skor"] = 310
    df["hadir"] = "TRUE"
    progress_bar["value"] = 80
    df = df.sort_values(by=["ID"]).reset_index(drop=True)

    # ubah Dtype
    df["ID"] = df["ID"].astype("int64")
    df["l1"] = df["l1"].astype("int64")
    df["g1"] = df["g1"].astype("int64")
    df["r1"] = df["r1"].astype("int64")
    df["l2"] = df["l2"].astype("int64")
    df["g2"] = df["g2"].astype("int64")
    df["r2"] = df["r2"].astype("int64")
    df["skor"] = df["skor"].astype("int64")
    progress_bar["value"] = 90
    # save table to same path as first file
    output_path = os.path.join(os.path.dirname(file_paths[0]), "output.xlsx")
    df.to_excel(output_path, index=False)
    progress_bar["value"] = 100
    progress_bar.stop()
    messagebox.showinfo("Success", "Table saved successfully!")
    output_path = os.path.join(os.path.dirname(file_paths[0]), "output.xlsx")
    os.startfile(os.path.dirname(output_path))

    root.destroy()


# Create the main window
root = ctk.CTk()
root.title("Select Multiple Excel Files")
width = 300
height = 300
x = (root.winfo_screenwidth() - width) // 2
y = (root.winfo_screenheight() - height) // 2
root.geometry(f"{width}x{height}+{x}+{y}")
# Create the button to select files
button = ctk.CTkButton(root, text="Select Files", command=select_files)
button.pack(pady=50)  # Set the button's y position to 50 pixels

frame = ttk.Frame(root, padding=10)
frame.pack(fill=tk.BOTH, expand=True)


progress_bar = ttk.Progressbar(
    frame, orient="horizontal", length=200, mode="determinate"
)
progress_bar.pack(pady=10)

if __name__ == "__main__":
    # Run the main loop
    root.mainloop()
