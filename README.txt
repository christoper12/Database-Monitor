========================================================
MySQL Database Size Monitor
========================================================

Author  : CSangjaya
Version : 1.0
Tech    : Python + CustomTkinter + MySQL
Purpose : Real-time MySQL database monitoring tool

========================================================
FEATURES
========================================================

✓ Live MySQL database monitoring
✓ Auto refresh every configurable interval
✓ Dynamic countdown timer
✓ Search database/table
✓ Database filter dropdown
✓ Sortable columns
✓ Export to Excel
✓ Top 10 biggest tables
✓ Biggest database summary
✓ Biggest table summary
✓ Total database size
✓ Database growth tracking
✓ JSON snapshot history
✓ Progress bar loading
✓ Responsive UI
✓ Scrollable summary panels

========================================================
DISPLAYED COLUMNS
========================================================

- Database
- Table
- Size MB
- Engine
- Rows
- Data MB
- Index MB

========================================================
GROWTH HISTORY SYSTEM
========================================================

Application stores database size snapshots into:

    history.json

Snapshots are saved only when:

✓ Manual refresh
✓ Database size changed
✓ More than 1 hour passed

This prevents history.json from growing too fast.

========================================================
AUTO REFRESH
========================================================

Default auto refresh interval:

    5 minutes

Countdown display example:

    Next refresh in: 04:59

Auto refresh can be enabled/disabled from checkbox.

========================================================
REQUIREMENTS
========================================================

Python 3.11+

Required packages:

pip install customtkinter
pip install pandas
pip install sqlalchemy
pip install pymysql
pip install openpyxl

Or:

pip install -r requirements.txt

========================================================
CONFIGURATION
========================================================

Create config.ini file:

--------------------------------------------------------

[MYSQL]
HOST=localhost
USER=root
PASSWORD=yourpassword
PORT=3306

--------------------------------------------------------

========================================================
RUN APPLICATION
========================================================

python app.py

========================================================
EXPORT EXCEL
========================================================

Export button generates Excel report containing:

✓ Database name
✓ Table name
✓ Size MB
✓ Engine
✓ Rows
✓ Data Size
✓ Index Size

========================================================
OBSERVABILITY FEATURES
========================================================

This application acts like a lightweight MySQL observability tool.

Useful for:

✓ Detecting abnormal database growth
✓ Detecting bloated indexes
✓ Monitoring table growth
✓ Capacity planning
✓ Identifying largest tables
✓ Tracking storage usage trends

========================================================
INDEX BLOAT DETECTION
========================================================

If:

- Size MB is very large
- Rows count is small
- Index MB is huge

Then indexes may be bloated.

Example:

Large table size
+
Few rows
+
Huge index size

= Possible index optimization needed

========================================================
FILES
========================================================

app.py
    Main application

config.ini
    MySQL configuration

history.json
    Database growth snapshots

exportDatabaseSizeExcel.py
    Excel export module

========================================================
NOTES
========================================================

- TABLE_ROWS from MySQL may be estimated
  depending on storage engine.

- InnoDB usually provides approximate row counts.

- Large databases may take several seconds
  during refresh.

========================================================
END
========================================================