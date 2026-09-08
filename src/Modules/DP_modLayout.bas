Attribute VB_Name = "DP_modLayout"
Option Explicit

'----------------------------------------
' Picker
'----------------------------------------

Public Const DP_DATEPICKER_WIDTH As Long = 236

Public Const DP_SHOW_TITLEBAR As Boolean = False

Private Const DATEPICKER_MARGIN_RATIO As Single = 0.05
Private Const DATEPICKER_MARGIN As Single = _
    DP_DATEPICKER_WIDTH * DATEPICKER_MARGIN_RATIO

Public Const DP_MONTHPICKER_WIDTH As Single = _
    DP_DATEPICKER_WIDTH

Public Const DP_YEARPICKER_WIDTH As Single = _
    DP_DATEPICKER_WIDTH

Private Const SECTION_GAP_RATIO As Single = 0.11


'----------------------------------------
' Day
'----------------------------------------

Public Const DP_CALENDAR_COLUMNS As Long = 7
Public Const DP_CALENDAR_ROWS As Long = 6

Private Const GRID_WIDTH As Single = _
    DP_DATEPICKER_WIDTH - DATEPICKER_MARGIN * 2

Public Const DP_GRID_CELL_WIDTH As Single = _
    GRID_WIDTH / DP_CALENDAR_COLUMNS

Public Const DP_GRID_CELL_HEIGHT As Single = _
    DP_GRID_CELL_WIDTH

Private Const GRID_HEIGHT As Single = _
    DP_GRID_CELL_HEIGHT * (DP_CALENDAR_ROWS + 1)

Public Const DP_GRID_LEFT As Single = _
    DATEPICKER_MARGIN


Public Const DP_CALENDAR_CELL_WIDTH As Single = _
    DP_GRID_CELL_WIDTH * 0.92

Public Const DP_CALENDAR_CELL_HEIGHT As Single = _
    DP_CALENDAR_CELL_WIDTH

Private Const CALENDAR_SECTION_GAP As Single = _
    DP_GRID_CELL_HEIGHT * SECTION_GAP_RATIO


'----------------------------------------
' Period (Month or Year)
'----------------------------------------

Public Const DP_PERIOD_GRID_CELL_HEIGHT As Single = _
    DP_GRID_CELL_HEIGHT

Public Const DP_PERIOD_CELL_HEIGHT As Single = _
    DP_PERIOD_GRID_CELL_HEIGHT * 0.92

Public Const DP_PERIOD_GRID_LEFT As Single = _
    DATEPICKER_MARGIN

Private Const PERIOD_SECTION_GAP As Single = _
    DP_PERIOD_GRID_CELL_HEIGHT * SECTION_GAP_RATIO


'----------------------------------------
' Month
'----------------------------------------

Public Const DP_MONTH_GRID_COLUMNS As Long = 3
Public Const DP_MONTH_GRID_ROWS As Long = 4

Private Const MONTH_GRID_WIDTH As Single = _
    DP_MONTHPICKER_WIDTH - DATEPICKER_MARGIN * 2

Public Const DP_MONTH_GRID_CELL_WIDTH As Single = _
    MONTH_GRID_WIDTH / DP_MONTH_GRID_COLUMNS

Public Const DP_MONTH_CELL_WIDTH As Single = _
    DP_MONTH_GRID_CELL_WIDTH * 0.96

Private Const MONTH_HEIGHT As Single = _
    DP_PERIOD_GRID_CELL_HEIGHT * DP_MONTH_GRID_ROWS


'----------------------------------------
' Year
'----------------------------------------

Public Const DP_YEAR_GRID_COLUMNS As Long = 4
Public Const DP_YEAR_GRID_ROWS As Long = 4

Private Const YEAR_GRID_WIDTH As Single = _
    DP_YEARPICKER_WIDTH - DATEPICKER_MARGIN * 2

Public Const DP_YEAR_GRID_CELL_WIDTH As Single = _
    YEAR_GRID_WIDTH / DP_YEAR_GRID_COLUMNS

Public Const DP_YEAR_CELL_WIDTH As Single = _
    DP_YEAR_GRID_CELL_WIDTH * 0.94

Private Const YEAR_HEIGHT As Single = _
    DP_PERIOD_GRID_CELL_HEIGHT * DP_YEAR_GRID_ROWS


'----------------------------------------
' Header
'----------------------------------------

Public Const DP_HEADER_TOP As Single = _
    DATEPICKER_MARGIN

Public Const DP_HEADER_WIDTH As Single = _
    DP_GRID_CELL_WIDTH * 5 * 0.98

Public Const DP_HEADER_HEIGHT As Single = _
    DP_GRID_CELL_HEIGHT * 0.7

Public Const DP_HEADER_LEFT As Single = _
    DP_GRID_LEFT


Public Const DP_ARROW_HEIGHT As Single = _
    DP_HEADER_HEIGHT

Public Const DP_ARROW_WIDTH As Single = _
    DP_CALENDAR_CELL_WIDTH

Private Const CELL_RIGHT_OFFSET As Single = _
    (DP_GRID_CELL_WIDTH - DP_CALENDAR_CELL_WIDTH) * 0.5

Public Const DP_ARROW_PREV_LEFT As Single = _
    DP_GRID_LEFT + 5 * DP_GRID_CELL_WIDTH - CELL_RIGHT_OFFSET

Public Const DP_ARROW_NEXT_LEFT As Single = _
    DP_GRID_LEFT + 6 * DP_GRID_CELL_WIDTH - CELL_RIGHT_OFFSET

Public Const DP_ARROW_TOP As Single = _
    DP_HEADER_TOP


'----------------------------------------
' Footer
'----------------------------------------

Public Const DP_ACTION_BUTTON_HEIGHT As Single = _
    DP_GRID_CELL_HEIGHT * 0.7

Public Const DP_ACTION_BUTTON_BOTTOM_MARGIN As Single = _
    DATEPICKER_MARGIN

Public Const DP_ACTION_BUTTON_RIGHT_MARGIN As Single = _
    DATEPICKER_MARGIN


'----------------------------------------
' Positioning
'----------------------------------------

Public Const DP_GRID_TOP As Single = _
    DP_HEADER_TOP + _
    DP_HEADER_HEIGHT + _
    CALENDAR_SECTION_GAP

Public Const DP_PERIOD_GRID_TOP As Single = _
    DP_HEADER_TOP + _
    DP_HEADER_HEIGHT + _
    PERIOD_SECTION_GAP


Private Const PICKER_FIXED_HEIGHT As Single = _
    DATEPICKER_MARGIN * 2 + _
    DP_HEADER_HEIGHT + _
    DP_ACTION_BUTTON_HEIGHT

Public Const DP_DATEPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    CALENDAR_SECTION_GAP * 2 + _
    GRID_HEIGHT

Public Const DP_MONTHPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    PERIOD_SECTION_GAP * 2 + _
    MONTH_HEIGHT

Public Const DP_YEARPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    PERIOD_SECTION_GAP * 2 + _
    YEAR_HEIGHT


'----------------------------------------
' Responsive scaling
'----------------------------------------

' Reference width for responsive scaling.
' Must remain 236. This is the original design width.
Public Const DP_DATEPICKER_BASE_WIDTH As Long = 236

