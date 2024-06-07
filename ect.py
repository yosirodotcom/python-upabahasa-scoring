import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import warnings
import os

# import missingno as msno


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
            62,
            63,
            64,
            65,
            66,
            67,
            68,
            69,
            70,
            71,
            72,
            73,
            74,
            75,
            76,
            77,
            78,
            79,
            80,
            81,
            82,
            83,
            84,
            85,
            86,
            87,
            88,
            89,
            90,
            91,
            92,
            93,
            94,
            95,
            96,
            97,
            98,
            99,
            100,
        ],
        "l2": [
            25,
            5,
            5,
            5,
            5,
            5,
            5,
            10,
            15,
            20,
            25,
            30,
            35,
            40,
            45,
            50,
            55,
            60,
            65,
            70,
            75,
            80,
            85,
            90,
            95,
            100,
            110,
            115,
            120,
            125,
            130,
            135,
            140,
            145,
            150,
            160,
            165,
            170,
            175,
            180,
            185,
            190,
            195,
            200,
            210,
            215,
            220,
            230,
            240,
            245,
            250,
            255,
            260,
            270,
            275,
            280,
            290,
            295,
            300,
            310,
            315,
            320,
            325,
            330,
            340,
            345,
            350,
            360,
            365,
            370,
            380,
            385,
            390,
            395,
            400,
            405,
            410,
            420,
            425,
            430,
            440,
            445,
            450,
            460,
            465,
            470,
            475,
            480,
            485,
            490,
            495,
            495,
            495,
            495,
            495,
            495,
            495,
            495,
            495,
            495,
            495,
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
            62,
            63,
            64,
            65,
            66,
            67,
            68,
            69,
            70,
            71,
            72,
            73,
            74,
            75,
            76,
            77,
            78,
            79,
            80,
            81,
            82,
            83,
            84,
            85,
            86,
            87,
            88,
            89,
            90,
            91,
            92,
            93,
            94,
            95,
            96,
            97,
            98,
            99,
            100,
        ],
        "r2": [
            22,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            5,
            10,
            15,
            20,
            25,
            30,
            35,
            40,
            45,
            50,
            60,
            65,
            70,
            80,
            85,
            90,
            95,
            100,
            110,
            115,
            120,
            125,
            130,
            140,
            145,
            150,
            160,
            165,
            170,
            175,
            180,
            190,
            195,
            200,
            210,
            215,
            220,
            225,
            230,
            235,
            240,
            250,
            255,
            260,
            265,
            270,
            280,
            285,
            290,
            300,
            305,
            310,
            320,
            325,
            330,
            335,
            340,
            350,
            355,
            360,
            365,
            370,
            380,
            385,
            390,
            395,
            400,
            405,
            410,
            415,
            420,
            425,
            430,
            435,
            445,
            450,
            455,
            465,
            470,
            480,
            485,
            490,
            495,
            495,
            495,
            495,
        ],
    }
)


def select_files():
    progress_bar.start()
    filenames = filedialog.askopenfilenames(
        title="Select Excel Files", filetypes=[("Excel Files", "*.xls")]
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
    df = df[["No. Peserta", "Section 1", "Section 2"]]
    df.columns = ["ID", "l1", "r1"]
    progress_bar["value"] = 60
    df = df[(df["l1"] != 0) | (df["r1"] != 0)]
    # print(df.info())
    df = df.merge(rumus_listening, on="l1", how="left")
    df = df.merge(rumus_reading, on="r1", how="left")
    progress_bar["value"] = 70
    df["skor"] = round((df["l2"] + df["r2"]))
    df.loc[df["skor"] < 10, "skor"] = 10
    df["hadir"] = "TRUE"
    progress_bar["value"] = 80
    df = df.sort_values(by=["ID"]).reset_index(drop=True)

    # ubah Dtype
    df["ID"] = df["ID"].astype("int64")
    df["l1"] = df["l1"].astype("int64")

    df["r1"] = df["r1"].astype("int64")
    df["l2"] = df["l2"].astype("int64")

    df["r2"] = df["r2"].astype("int64")
    df["skor"] = df["skor"].astype("int64")
    progress_bar["value"] = 90
    # save table to same path as first file
    output_path = os.path.join(os.path.dirname(file_paths[0]), "output.xlsx")
    df.to_excel(output_path, index=False)
    progress_bar["value"] = 100
    progress_bar.stop()
    messagebox.showinfo("Success", "Table saved successfully!")
    output_path = os.path.join(os.path.dirname(file_paths[0]), "output_ect.xlsx")
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
