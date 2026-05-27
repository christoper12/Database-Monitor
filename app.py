import customtkinter as ctk
import tkinter.ttk as ttk
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime
import configparser
from exportDatabaseSizeExcel import export_excel_file
import json
import os
import subprocess

# =========================
# LOAD CONFIG
# =========================
config = configparser.ConfigParser()
config.read('config.ini')

HOST = config['MYSQL']['HOST']
USER = config['MYSQL']['USER']
PASSWORD = config['MYSQL']['PASSWORD']
PORT = config['MYSQL']['PORT']

# =========================
# MYSQL ENGINE
# =========================
engine = create_engine(
    f"mysql+pymysql://{USER}:{PASSWORD}@{HOST}:{PORT}"
)

# =========================
# APP CONFIG
# =========================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("MySQL Database Size Monitor")
# =========================
# WINDOW SIZE
# =========================
window_width = 1200
window_height = 850
app.geometry(f"{window_width}x{window_height}")

# Get screen size
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()

# Calculate center position
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))

# Set window position
app.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Minimum size
app.minsize(1200, 850)


# =========================
# TREEVIEW STYLE
# =========================
style = ttk.Style()

style.theme_use("default")

style.configure(
    "Treeview",
    background="#838383",
    foreground="white",
    fieldbackground="#838383",
    bordercolor="#616569",
    borderwidth=5,
    rowheight=38,
    relief="flat",
    font=("Arial", 12)
)

style.layout(
    "Treeview",
    [('Treeview.treearea', {'sticky': 'nswe'})]
)

style.map(
    'Treeview',
    background=[('selected', '#22559b')]
)

style.configure(
    "Treeview.Heading",
    background="#1f1f1f",
    foreground="white",
    relief="flat",
    padding=(10, 12),
    font=("Arial", 13, "bold")
)

style.map(
    "Treeview.Heading",
    background=[('active', "#686767")]
)

# =========================
# HEADER
# =========================
header = ctk.CTkLabel(
    app,
    text="MySQL Database Size Monitor",
    font=("Arial", 20, "bold")
)
header.pack(pady=20)

# =========================
# BUTTON FRAME
# =========================
button_frame = ctk.CTkFrame(app)
button_frame.pack(fill='x', padx=20, pady=10)

# =========================
# SEARCH ENTRY
# =========================
search_entry = ctk.CTkEntry(
    button_frame,
    placeholder_text="Search database or table..."
)

search_entry.pack(
    side='right',
    padx=10,
    pady=10
)

# =========================
# TABLE FRAME
# =========================
table_frame = ctk.CTkFrame(app)
table_frame.pack(fill='both', expand=True, padx=20, pady=10)

# =========================
# TREEVIEW
# =========================
columns = ('Database', 'Table', 'Total Size', 'Engine', 'Rows', 'Data Size', 'Index Size')

my_table = ttk.Treeview(
    table_frame,
    columns=columns,
    show='headings'
)

for col in columns:

    my_table.heading(
        col,
        text=col,
        command=lambda c=col: sort_column(my_table, c)
    )

    my_table.column(col, width=250)

    if col in ('Rows', 'Data Size', 'Index Size', 'Total Size'):
        my_table.column(col, anchor='e', width=120)
    else:
        my_table.column(col, anchor='w', width=220)


# =========================
# SCROLLBAR
# =========================
scrollbar = ttk.Scrollbar(
    table_frame,
    orient='vertical',
    command=my_table.yview
)

my_table.configure(yscroll=scrollbar.set)

# =========================
# PACK
# =========================
scrollbar.pack(side='right', fill='y')

my_table.pack(
    side='left',
    fill='both',
    expand=True
)

# =========================
# STATUS LABEL
# =========================
status_label = ctk.CTkLabel(
    app,
    text='Status: Ready'
)
status_label.pack(pady=10)

# =========================
# COUNTDOWN LABEL
# =========================
countdown_label = ctk.CTkLabel(
    app,
    text="Next refresh in: 30 sec"
)

countdown_label.pack()

# =========================
# STATS CONTAINER
# =========================
stats_frame = ctk.CTkFrame(
    app,
    fg_color="transparent"
)

stats_frame.pack(
    fill='x',
    padx=20,
    pady=10
)

# Responsive columns
stats_frame.grid_columnconfigure(0, weight=1)
stats_frame.grid_columnconfigure(1, weight=1)

# =========================
# SUMMARY FRAME
# =========================
summary_frame = ctk.CTkScrollableFrame(
    stats_frame,
    height=220
)

summary_frame.grid(
    row=0,
    column=0,
    sticky='nsew',
    padx=(0, 10)
)

# =========================
# TOP 10 FRAME
# =========================
top10_frame = ctk.CTkScrollableFrame(
    stats_frame,
    height=220
)

top10_frame.grid(
    row=0,
    column=1,
    sticky='nsew',
    padx=(10, 0)
)

# =========================
# SUMMARY LABEL
# =========================
summary_label = ctk.CTkLabel(
    summary_frame,
    text="Loading summary...",
    justify='left',
    anchor='w',
    font=("Arial", 14, "bold"),
    wraplength=500
)

summary_label.pack(
    anchor='w',
    padx=15,
    pady=15
)

# =========================
# TOP 10 TITLE
# =========================
top10_title = ctk.CTkLabel(
    top10_frame,
    text="Top 10 Biggest Tables",
    font=("Arial", 16, "bold")
)

top10_title.pack(
    anchor='w',
    padx=10,
    pady=(10, 5)
)

# =========================
# TOP 10 LABEL
# =========================
top10_label = ctk.CTkLabel(
    top10_frame,
    text="",
    justify='left',
    anchor='w',
    font=("Consolas", 12),
    wraplength=500
)

top10_label.pack(
    anchor='w',
    padx=10,
    pady=(0, 10)
)

# Equal height
stats_frame.grid_rowconfigure(0, weight=1)

# =========================
# PROGRESS BAR
# =========================
progress_bar = ctk.CTkProgressBar(app)

progress_bar.pack(
    fill='x',
    padx=20,
    pady=(0, 20)
)

progress_bar.set(0)

# =========================
# SORT STATES
# =========================
sort_states = {}

# =========================
# FORMAT SIZE
# =========================
def format_size(size_mb):

    if size_mb >= 1024 * 1024:
        return f"{size_mb / 1024 / 1024:.2f} TB"

    elif size_mb >= 1024:
        return f"{size_mb / 1024:.2f} GB"

    return f"{size_mb:.2f} MB"

# =========================
# SORT TREEVIEW COLUMN
# =========================
def sort_column(tree, col):

    reverse = sort_states.get(col, False)

    data_list = []

    for child in tree.get_children(''):

        value = tree.set(child, col)

        # Numeric conversion
        if col == 'Size MB':
            try:
                value = float(value)
            except:
                value = 0

        data_list.append((value, child))

    # Sort
    data_list.sort(reverse=reverse)

    # Rearrange rows
    for index, (_, child) in enumerate(data_list):
        tree.move(child, '', index)

    # Reset all headers
    for c in columns:
        my_table.heading(
            c,
            text=c,
            command=lambda _c=c: sort_column(my_table, _c)
        )

    # Add arrow icon
    arrow = "▼" if reverse else "▲"

    my_table.heading(
        col,
        text=f"{col} {arrow}",
        command=lambda: sort_column(my_table, col)
    )

    # Toggle next state
    sort_states[col] = not reverse

# =========================
# GLOBAL DATAFRAME
# =========================
global_df = pd.DataFrame()

# =========================
# POPULATE TABLE
# =========================
def populate_table(df):

    # Clear table
    for item in my_table.get_children():
        my_table.delete(item)

    # Insert rows
    for index, row in df.iterrows():

        # Zebra rows
        if index % 2 == 0:
            tag = 'evenrow'
        else:
            tag = 'oddrow'

        my_table.insert(
            '',
            'end',
            values=(
                row['Database'],
                row['Table'],
                format_size(row['Total Size']),
                row['Engine'],
                row['Rows'],
                format_size(row['Data Size']),
                format_size(row['Index Size']),
            ),
            tags=(tag,)
        )

# =========================
# DATABASE FILTER
# =========================
def filter_database(choice):

    global selected_database

    selected_database = choice

    apply_filters()

# =========================
# UPDATE STATUS
# =========================
# =========================
# UPDATE STATUS
# =========================
def update_status(df):

    now = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

    total_tables = len(df)
    total_databases = df['Database'].nunique()

    # Total size MB
    total_size_mb = df['Total Size'].sum()

    # Convert to GB if large
    total_size = format_size(total_size_mb)

    status_label.configure(
        text=(
            f'Databases: {total_databases} | '
            f'Tables: {total_tables} | '
            f'Total Size: {total_size} | '
            f'Last Refresh: {now}'
        )
    )

# =========================
# APPLY FILTERS
# =========================
def apply_filters():

    keyword = search_entry.get().lower()

    filtered_df = global_df.copy()

    # Filter database
    if selected_database != "ALL":

        filtered_df = filtered_df[
            filtered_df['Database'] == selected_database
        ]

    # Filter keyword
    if keyword != "":

        filtered_df = filtered_df[
            filtered_df['Database'].str.lower().str.contains(keyword)
            |
            filtered_df['Table'].str.lower().str.contains(keyword)
        ]

    populate_table(filtered_df)
    update_status(filtered_df)
    update_summary(filtered_df)
    update_top10_tables(filtered_df)

# =========================
# REALTIME SEARCH
# =========================
def realtime_search(event=None):

    apply_filters()


# =========================
# EXPORT EXCEL
# =========================
def export_excel():
    status_label.configure(
        text='Status: Exporting Excel...'
    )

    progress_bar.set(0.2)
    app.update()

    try:
        output_file = export_excel_file()

        progress_bar.set(1)
        app.update()

        status_label.configure(
            text=f'Export Success: {output_file}'
        )

        app.after(
            1000,
            lambda: progress_bar.set(0)
        )

    except Exception as e:
        status_label.configure(
            text=f'Export Error: {e}'
        )

        app.after(
            1000,
            lambda: progress_bar.set(0)
        )

# =========================
# OPEN MASTER EXCEL
# =========================
def open_master_excel():

    excel_file = "Master Excel Database Live 22.xlsx"

    if not os.path.exists(excel_file):

        status_label.configure(
            text=f"File not found: {excel_file}"
        )

        return

    try:

        os.startfile(excel_file)

        status_label.configure(
            text=f"Opened: {excel_file}"
        )

    except Exception as e:

        status_label.configure(
            text=f"Open Excel Error: {e}"
        )

# =========================
# AUTO REFRESH CONFIG
# =========================
AUTO_REFRESH_MS = 300000  # 5 minutes

# Save history every 1 hour
HISTORY_INTERVAL_SECONDS = 3600

auto_refresh_enabled = ctk.BooleanVar(value=True)
countdown_seconds = AUTO_REFRESH_MS // 1000

# =========================
# COUNTDOWN TIMER
# =========================
def update_countdown():

    global countdown_seconds

    if auto_refresh_enabled.get():

        # Convert seconds to MM:SS
        minutes = countdown_seconds // 60
        seconds = countdown_seconds % 60

        countdown_label.configure(
            text=f"Next refresh in: {minutes:02}:{seconds:02}"
        )

        countdown_seconds -= 1

        if countdown_seconds < 0:

            countdown_seconds = AUTO_REFRESH_MS // 1000

    else:

        countdown_label.configure(
            text="Auto Refresh Disabled"
        )

    app.after(
        1000,
        update_countdown
    )

# =========================
# AUTO REFRESH
# =========================
def auto_refresh():

    global countdown_seconds

    if auto_refresh_enabled.get():

        print("AUTO REFRESH RUNNING...")

        load_data()

        countdown_seconds = AUTO_REFRESH_MS // 1000

    app.after(
        AUTO_REFRESH_MS,
        auto_refresh
    )

# =========================
# UPDATE SUMMARY
# =========================
def update_summary(df):

    if df.empty:

        summary_label.configure(
            text="No data"
        )

        return

    # Biggest database
    grouped = (
        df.groupby('Database')['Total Size']
        .sum()
        .sort_values(ascending=False)
    )

    biggest_db = grouped.index[0]
    biggest_db_size = grouped.iloc[0]

    # Biggest table
    largest_row = df.sort_values(
        by='Total Size',
        ascending=False
    ).iloc[0]

    biggest_table = (
        f"{largest_row['Database']}."
        f"{largest_row['Table']}"
    )

    biggest_table_size = largest_row['Total Size']

    # Total size
    total_size_mb = df['Total Size'].sum()

    # Convert sizes
    def format_size(size_mb):

        if size_mb >= 1024:
            return f"{size_mb / 1024:.2f} GB"

        return f"{size_mb:.2f} MB"

    growth_text = show_growth()

    summary_text = (
        f"Biggest Database : "
        f"{biggest_db} "
        f"({format_size(biggest_db_size)})\n\n"

        f"Biggest Table    : "
        f"{biggest_table} "
        f"({format_size(biggest_table_size)})\n\n"

        f"Total Size       : "
        f"{format_size(total_size_mb)}"

        f"\n\nRecent Growth:\n{growth_text}"
    )

    summary_label.configure(
        text=summary_text
    )


# =========================
# TOP 10 BIGGEST TABLES
# =========================
def update_top10_tables(df):

    if df.empty:

        top10_label.configure(
            text="No data"
        )

        return

    # Top 10 largest
    top10 = df.sort_values(
        by='Total Size',
        ascending=False
    ).head(10)

    lines = []

    for index, row_data in enumerate(
        top10.iterrows(),
        start=1
    ):

        _, row = row_data

        size_mb = row['Total Size']

        size_text = format_size(size_mb)

        line = (
            f"{index:>2}. "
            f"{row['Database']}.{row['Table']} "
            f"({size_text})"
        )

        lines.append(line)

    top10_label.configure(
        text="\n".join(lines)
    )

# =========================
# SAVE HISTORY
# =========================
def save_history(df, force=False):
    from datetime import datetime

    history_file = "history.json"

    # Current grouped size
    grouped = (
        df.groupby('Database')['Total Size']
        .sum()
        .to_dict()
    )

    current_time = datetime.now()

    snapshot = {
        "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
        "databases": grouped
    }

    history = []

    # =========================
    # LOAD EXISTING HISTORY
    # =========================
    if os.path.exists(history_file):

        try:

            with open(history_file, "r") as f:
                history = json.load(f)

        except:
            history = []

    # =========================
    # FIRST SNAPSHOT
    # =========================
    if not history:

        history.append(snapshot)

        with open(history_file, "w") as f:
            json.dump(history, f, indent=4)

        # print("Initial history snapshot saved")

        return

    # =========================
    # GET LAST SNAPSHOT
    # =========================
    last_snapshot = history[-1]

    last_time = datetime.strptime(
        last_snapshot['timestamp'],
        "%Y-%m-%d %H:%M:%S"
    )

    seconds_diff = (
        current_time - last_time
    ).total_seconds()

    # =========================
    # CHECK SIZE CHANGES
    # =========================
    last_databases = last_snapshot['databases']

    changed = False

    for db_name, current_size in grouped.items():

        old_size = last_databases.get(db_name, 0)

        # tolerance 1 MB
        if abs(current_size - old_size) >= 1:

            changed = True
            break

    # =========================
    # SAVE CONDITIONS
    # =========================
    should_save = (
        force
        or changed
        or seconds_diff >= HISTORY_INTERVAL_SECONDS
    )

    if not should_save:

        print("Skip history save (no significant changes)")
        return

    # =========================
    # SAVE SNAPSHOT
    # =========================
    history.append(snapshot)

    # Keep latest 500 snapshots
    history = history[-500:]

    with open(history_file, "w") as f:
        json.dump(history, f, indent=4)

    # print("History snapshot saved")

# =========================
# SHOW GROWTH
# =========================
def show_growth():

    import json
    import os

    history_file = "history.json"

    if not os.path.exists(history_file):
        return "No history yet"

    if os.path.getsize(history_file) == 0:
        return "No history yet"

    try:

        with open(history_file, "r") as f:
            history = json.load(f)

    except Exception as e:

        print(e)

        return "History corrupted"

    # Minimal 2 snapshot
    if len(history) < 2:
        return "Not enough history"

    latest = history[-1]
    previous = history[-2]

    latest_db = latest['databases']
    previous_db = previous['databases']

    growth_data = []

    for db_name, latest_size in latest_db.items():

        old_size = previous_db.get(db_name, 0)

        growth = latest_size - old_size

        # Skip jika tidak berubah
        if growth == 0:
            continue

        growth_data.append(
            (db_name, growth)
        )

    # Tidak ada perubahan
    if not growth_data:
        return "No database growth detected"

    # Sort biggest growth first
    growth_data.sort(
        key=lambda x: x[1],
        reverse=True
    )

    lines = []

    for db_name, growth in growth_data:

        # Positive growth
        if growth > 0:

            growth_text = f"+{format_size(growth)}"

        # Negative growth
        else:

            growth = abs(growth)

            growth_text = f"-{format_size(growth)}"

        lines.append(
            f"{db_name} {growth_text}"
        )

    return "\n".join(lines)

# =========================
# LOAD DATA FUNCTION
# =========================
selected_database = "ALL"

def load_data(force_history=False):

    status_label.configure(text='Status: Loading data...')
    progress_bar.set(0.1)
    app.update()

    query = """
    SELECT
    TABLE_SCHEMA AS 'Database',
    TABLE_NAME AS 'Table',
    ENGINE AS 'Engine',
    TABLE_ROWS AS 'Rows',
    ROUND(DATA_LENGTH / 1024 / 1024, 2) AS 'Data Size',
    ROUND(INDEX_LENGTH / 1024 / 1024, 2) AS 'Index Size',
    (DATA_LENGTH + INDEX_LENGTH) / 1024 / 1024 AS 'Total Size'
FROM information_schema.TABLES
WHERE TABLE_SCHEMA NOT IN (
    'information_schema',
    'performance_schema',
    'mysql',
    'sys'
)
ORDER BY (DATA_LENGTH + INDEX_LENGTH) DESC;
    """

    try:
        df = pd.read_sql(query, engine)

    except Exception as e:

        status_label.configure(
            text=f"ERROR: {e}"
        )

        print(e)

        return

    global global_df
    global_df = df.copy()

    progress_bar.set(0.5)
    app.update()

    # Update database dropdown
    database_list = sorted(df['Database'].unique().tolist())

    database_dropdown.configure(
        values=["ALL"] + database_list
    )

    # Reset dropdown to ALL
    database_dropdown.set("ALL")

    # Reset selected database
    global selected_database
    selected_database = "ALL"

    populate_table(df)

    progress_bar.set(0.8)
    app.update()

    update_status(df)
    save_history(df,force=force_history)
    update_summary(df)
    update_top10_tables(df)

    progress_bar.set(1)
    app.update()

    app.after(
        500,
        lambda: progress_bar.set(0)
    )
# Realtime search
search_entry.bind('<KeyRelease>', realtime_search)

# =========================
# DATABASE DROPDOWN
# =========================
database_dropdown = ctk.CTkOptionMenu(
    button_frame,
    values=["ALL"],
    command=filter_database
)

database_dropdown.pack(
    side='right',
    padx=10
)

database_dropdown.set("ALL")

# =========================
# REFRESH BUTTON
# =========================
refresh_button = ctk.CTkButton(
    button_frame,
    text='Refresh',
    command=lambda: load_data(force_history=True)
)

refresh_button.pack(side='left', padx=10, pady=10)

# =========================
# EXPORT BUTTON
# =========================
export_button = ctk.CTkButton(
    button_frame,
    text='Export Excel',
    command=export_excel
)

export_button.pack(
    side='left',
    padx=10,
    pady=10
)

# =========================
# OPEN MASTER EXCEL BUTTON
# =========================
open_excel_button = ctk.CTkButton(
    button_frame,
    text='Open Master Excel',
    command=open_master_excel,
    fg_color="#1D6F42",
    hover_color="#14532d"
)

open_excel_button.pack(
    side='left',
    padx=10,
    pady=10
)

# =========================
# AUTO REFRESH CHECKBOX
# =========================
auto_refresh_checkbox = ctk.CTkCheckBox(
    button_frame,
    text="Auto Refresh",
    variable=auto_refresh_enabled
)

auto_refresh_checkbox.pack(
    side='left',
    padx=10
)

# =========================
# INITIAL LOAD
# =========================
load_data()
# Start countdown
update_countdown()

# Start auto refresh
app.after(
    AUTO_REFRESH_MS,
    auto_refresh
)

# =========================
# RUN APP
# =========================
app.mainloop()