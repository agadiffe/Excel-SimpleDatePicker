Attribute VB_Name = "DP_modLayout"
Option Explicit

'----------------------------------------
' Responsive scaling
'----------------------------------------

' Reference width for responsive scaling.
' Must remain 236. This is the original design width.
Public Const DATEPICKER_BASE_WIDTH As Long = 236


'----------------------------------------
' Picker
'----------------------------------------

Public Const DATEPICKER_WIDTH As Long = 236

Private Const DATEPICKER_MARGIN_RATIO As Single = 0.05
Private Const DATEPICKER_MARGIN As Single = _
    DATEPICKER_WIDTH * DATEPICKER_MARGIN_RATIO

Public Const MONTHPICKER_WIDTH As Single = _
    DATEPICKER_WIDTH

Public Const YEARPICKER_WIDTH As Single = _
    DATEPICKER_WIDTH

Private Const SECTION_GAP_RATIO As Single = 0.1

'----------------------------------------
' Day
'----------------------------------------

Public Const CALENDAR_COLUMNS As Long = 7
Public Const CALENDAR_ROWS As Long = 6

Public Const GRID_WIDTH As Single = _
    DATEPICKER_WIDTH * (1 - DATEPICKER_MARGIN_RATIO * 2)

Public Const GRID_CELL_WIDTH As Single = _
    GRID_WIDTH / CALENDAR_COLUMNS

Public Const GRID_CELL_HEIGHT As Single = _
    GRID_CELL_WIDTH

Public Const GRID_HEIGHT As Single = _
    GRID_CELL_HEIGHT * (CALENDAR_ROWS + 1)

Public Const GRID_LEFT As Single = _
    DATEPICKER_MARGIN


Public Const CALENDAR_CELL_WIDTH As Single = _
    GRID_CELL_WIDTH * 0.92

Public Const CALENDAR_CELL_HEIGHT As Single = _
    CALENDAR_CELL_WIDTH

Public Const CALENDAR_SECTION_GAP As Single = _
    GRID_CELL_HEIGHT * SECTION_GAP_RATIO


'----------------------------------------
' Period (Month or Year)
'----------------------------------------

Public Const PERIOD_GRID_CELL_HEIGHT As Single = _
    GRID_CELL_HEIGHT

Public Const PERIOD_CELL_HEIGHT As Single = _
    PERIOD_GRID_CELL_HEIGHT * 0.92

Public Const PERIOD_GRID_LEFT As Single = _
    DATEPICKER_MARGIN

Public Const PERIOD_SECTION_GAP As Single = _
    PERIOD_CELL_HEIGHT * SECTION_GAP_RATIO


'----------------------------------------
' Month
'----------------------------------------

Public Const MONTH_GRID_COLUMNS As Long = 3
Public Const MONTH_GRID_ROWS As Long = 4

Public Const MONTH_GRID_WIDTH As Single = _
    MONTHPICKER_WIDTH * (1 - DATEPICKER_MARGIN_RATIO * 2)

Public Const MONTH_GRID_CELL_WIDTH As Single = _
    MONTH_GRID_WIDTH / MONTH_GRID_COLUMNS

Public Const MONTH_CELL_WIDTH As Single = _
    MONTH_GRID_CELL_WIDTH * 0.96

Public Const MONTH_HEIGHT As Single = _
    PERIOD_GRID_CELL_HEIGHT * MONTH_GRID_ROWS


'----------------------------------------
' Year
'----------------------------------------

Public Const YEAR_GRID_COLUMNS As Long = 4
Public Const YEAR_GRID_ROWS As Long = 4

Public Const YEAR_BLOCK_SIZE As Long = _
    YEAR_GRID_COLUMNS * YEAR_GRID_ROWS

Public Const YEAR_GRID_WIDTH As Single = _
    YEARPICKER_WIDTH * (1 - DATEPICKER_MARGIN_RATIO * 2)

Public Const YEAR_GRID_CELL_WIDTH As Single = _
    YEAR_GRID_WIDTH / YEAR_GRID_COLUMNS

Public Const YEAR_CELL_WIDTH As Single = _
    YEAR_GRID_CELL_WIDTH * 0.94

Public Const YEAR_HEIGHT As Single = _
    PERIOD_GRID_CELL_HEIGHT * YEAR_GRID_ROWS


'----------------------------------------
' Header
'----------------------------------------

Public Const HEADER_TOP As Single = _
    DATEPICKER_MARGIN

Public Const HEADER_WIDTH As Single = _
    GRID_CELL_WIDTH * 5 * 0.98

Public Const HEADER_HEIGHT As Single = _
    GRID_CELL_HEIGHT * 0.7

Public Const HEADER_LEFT As Single = _
    GRID_LEFT


Public Const ARROW_HEIGHT As Single = _
    HEADER_HEIGHT

Public Const ARROW_WIDTH As Single = _
    CALENDAR_CELL_WIDTH

Private Const CELL_RIGHT_OFFSET As Single = _
    (GRID_CELL_WIDTH - CALENDAR_CELL_WIDTH) * 0.5

Public Const ARROW_PREV_LEFT As Single = _
    GRID_LEFT + 5 * GRID_CELL_WIDTH - CELL_RIGHT_OFFSET

Public Const ARROW_NEXT_LEFT As Single = _
    GRID_LEFT + 6 * GRID_CELL_WIDTH - CELL_RIGHT_OFFSET

Public Const ARROW_TOP As Single = _
    HEADER_TOP


'----------------------------------------
' Footer
'----------------------------------------

Public Const ACTION_BUTTON_HEIGHT As Single = _
    GRID_CELL_HEIGHT * 0.7

Public Const ACTION_BUTTON_BOTTOM_MARGIN As Single = _
    DATEPICKER_MARGIN

Public Const ACTION_BUTTON_RIGHT_MARGIN As Single = _
    DATEPICKER_MARGIN


'----------------------------------------
' Positioning
'----------------------------------------

Public Const GRID_TOP As Single = _
    HEADER_TOP + HEADER_HEIGHT + CALENDAR_SECTION_GAP

Public Const PERIOD_GRID_TOP As Single = _
    HEADER_TOP + HEADER_HEIGHT + PERIOD_SECTION_GAP


Private Const PICKER_FIXED_HEIGHT As Single = _
    DATEPICKER_MARGIN * 2 + _
    HEADER_HEIGHT + _
    ACTION_BUTTON_HEIGHT

Public Const DATEPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    CALENDAR_SECTION_GAP * 2 + _
    GRID_HEIGHT

Public Const MONTHPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    PERIOD_SECTION_GAP * 2 + _
    MONTH_HEIGHT

Public Const YEARPICKER_HEIGHT As Single = _
    PICKER_FIXED_HEIGHT + _
    PERIOD_SECTION_GAP * 2 + _
    YEAR_HEIGHT

