# Access Kanban Board — Architecture & Integration Guide

This document explains how the HTML/CSS/JavaScript frontend and VBA backend work together in the Access Kanban project organizer.

---

## Table of Contents

1. [High-Level Architecture](#high-level-architecture)
2. [Data Models](#data-models)
3. [VBA Backend Structure](#vba-backend-structure)
4. [Frontend Architecture](#frontend-architecture)
5. [Communication Bridge](#communication-bridge)
6. [Data Flow Examples](#data-flow-examples)
7. [Key Features](#key-features)

---

## High-Level Architecture

```text
┌──────────────────────────────────────────────────────────────┐
│                  ACCESS DATABASE (.accdb)                    │
│  ┌────────────────────────────────────────────────────────┐  │
│  │  Tables: tblBoards, tblTasks                           │  │
│  │  VBA Modules & Classes: ~1500 lines total              │  │
│  └────────────────────────────────────────────────────────┘  │
└────────────────┬─────────────────────────────────────────────┘
                 │ COM Bridge (clsHtmlBridge)
                 │ JSON Serialization
                 ↓
┌──────────────────────────────────────────────────────────────┐
│              WEBVIEW/BROWSER CONTROL (iframe)                │
│  ┌────────────────────────────────────────────────────────┐  │
│  │  index.html / kanban.html                              │  │
│  │  - ES6 Classes (KanbanApp, TaskEditor, etc.)           │  │
│  │  - Drag-drop board rendering                           │  │
│  │  - Modal dialogs for CRUD                              │  │
│  │  - Theme switching, live updates                       │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
```

**Key Flow:**

- User interacts with HTML UI (index.html rendered in browser control)
- JavaScript collects changes and calls `ParentBridge.persist(action, payload)`
- Payload is JSON-serialized and set as a "pending command"
- VBA form timer polls for pending commands via `clsHtmlBridge.GetPendingCommand()`
- Command is routed through `clsCommandRouter.Route()` to appropriate service
- Service updates database via Repository pattern
- Full data payload is rebuilt and pushed back to browser via `clsHtmlBridge.PushData()`

---

## Data Models

### Board Model

**VBA: `clsBoard` (Data Transfer Object)**

```vba
Public BoardID      As Long
Public BoardName    As String
Public BoardColor   As String
Public CreatedOn    As Variant      ' Date or Null
Public IsDeleted    As Boolean
```

**Key Methods:**

- `LoadFromRecordset(rs)` — Populate from database
- `LoadFromJson(jsonStr)` — Populate from JSON payload
- `ToJson()` — Serialize to JSON object string

**JSON Format:**

```json
{
  "id": 1,
  "name": "Q1 Sprint",
  "color": "#6c63ff"
}
```

---

### Task Model

**VBA: `clsTask` (Data Transfer Object)**

```vba
Public TaskID       As Long
Public BoardID      As Long
Public Title        As String
Public Description  As String
Public Priority     As String         ' "Low", "Medium", "High"
Public Status       As String         ' "backlog", "todo", "inprog", "review", "done"
Public StartDate    As Variant        ' Date or Null
Public EndDate      As Variant        ' Date or Null
Public StartTime    As Variant        ' Date or Null
Public EndTime      As Variant        ' Date or Null
Public PctComplete  As Integer        ' 0-100
Public TaskColor    As String         ' Hex color or ""
Public TaskType     As String         ' "Bug", "Feature", "Story", etc.
Public ParentID     As Variant        ' Long or Null (for subtasks)
Public Notes        As String
Public IsExpanded   As Boolean        ' For subtask visibility
Public SortOrder    As Long           ' Position within status column
Public IsDeleted    As Boolean        ' Soft-delete flag
```

**Key Methods:**

- `LoadFromRecordset(rs)` — Load from database row
- `LoadFromJson(jsonStr)` — Parse JSON payload
- `ToJson()` — Serialize to JSON object string

**JSON Format:**

```json
{
  "id": "t1234567890",
  "boardId": 1,
  "title": "Implement API",
  "desc": "Create REST endpoints",
  "notes": "Internal notes here",
  "priority": "High",
  "status": "inprog",
  "startDate": "2026-04-08",
  "endDate": "2026-04-15",
  "startTime": "09:00",
  "endTime": "17:00",
  "pctComplete": 45,
  "taskColor": "#ef4444",
  "taskType": "Feature",
  "parentId": "",
  "isExpanded": true,
  "sortOrder": 2
}
```

---

## VBA Backend Structure

### Architecture: Service + Repository Pattern

```text
┌─────────────────────────────────────────────┐
│        clsCommandRouter                     │
│  - Routes JSON commands to services         │
│  - Maintains ActiveBoardID state            │
└──────┬──────────────────┬───────────────────┘
       │                  │
       ↓                  ↓
┌─────────────────────┐ ┌─────────────────────┐
│ clsBoardService     │ │ clsTaskService      │
│ - AddBoard()        │ │ - AddTask()         │
│ - RenameBoard()     │ │ - UpdateTask()      │
│ - DeleteBoard()     │ │ - DeleteTask()      │
│ - UpdateColor()     │ │ - ToggleExpand()    │
└──────┬──────────────┘ └──────┬──────────────┘
       │                      │
       ↓                      ↓
┌─────────────────────┐ ┌─────────────────────┐
│ clsBoardRepo        │ │ clsTaskRepo         │
│ - GetAll()          │ │ - GetAll()          │
│ - GetById()         │ │ - GetById()         │
│ - Add()             │ │ - Add()             │
│ - Rename()          │ │ - Update()          │
│ - SoftDelete()      │ │ - SoftDelete()      │
└─────────┬───────────┘ └──────┬──────────────┘
          │                    │
          └────────┬──────────┘
                   ↓
          ┌─────────────────────┐
          │   DAO: CurrentDb    │
          │  tblBoards, tblTasks│
          └─────────────────────┘
```

### Class Hierarchy

#### `clsCommandRouter` — Command Dispatcher

Receives JSON commands from JavaScript and routes to appropriate service.

```vba
Public Function Route(ByVal jsonStr As String) As String
    Dim action As String
    action = ExtractJSONValue(jsonStr, "action")
    
    Select Case action
        Case "addBoard"
            BoardService.AddBoard jsonStr
            Route = ActiveBoardID
        Case "addTask"
            Route = TaskService.AddTask(jsonStr)
        Case "updateTask"
            Route = TaskService.UpdateTask(jsonStr)
        Case "deleteTask"
            Route = TaskService.DeleteTask(jsonStr)
        ' ... etc
    End Select
End Function
```

**Actions Handled:**

- `setActiveBoard` — Switch active board context
- `addBoard`, `renameBoard`, `updateBoard`, `deleteBoard`
- `addTask`, `updateTask`, `deleteTask`
- `toggleExpand` — Collapse/expand subtasks
- `reorderTask` — Drag-drop reordering

**Return Value:**

- Board ID to refresh, or `"NORELOAD"` to skip data fetch

---

#### `clsBoardService` & `clsTaskService` — Business Logic

Contains domain logic (validation, cascading deletes, etc.)

```vba
' clsBoardService example
Public Sub AddBoard(ByVal jsonStr As String)
    Dim b As New clsBoard
    b.LoadFromJson jsonStr
    BoardRepo.Add b
End Sub

' clsTaskService has more complex logic:
' - Cascade deletes children when parent task deleted
' - Auto-calculate parent % complete from subtasks
' - Auto-assign sort order based on column
```

---

#### `clsBoardRepo` & `clsTaskRepo` — Data Access

DAO-based record CRUD operations.

```vba
' clsTaskRepo.Add example
Public Sub Add(task As clsTask)
    ' Insert new row
    ' Auto-assign SortOrder = max(SortOrder) + 1 in column
    ' Set CreatedOn/ModifiedOn timestamps
End Sub

Public Sub SoftDelete(ByVal taskId As Long)
    ' Mark IsDeleted = True
    ' Cascade to children (ParentID = taskId)
End Sub
```

---

### JSON Utilities

#### `mod_JsonUtility` — JSON Helper Functions

```vba
Function JSONEscape(v As Variant) As String      ' Escape strings for JSON
Function ExtractJSONValue(jsonStr, key) As String ' Extract key value from flat JSON
Function BuildJsonPair(key, value, asString) As String ' "key":"value" or "key":num
Function WrapJsonObject(ParamArray pairs) As String    ' {...}
Function WrapJsonArray(items As Collection) As String  ' [...]
```

#### `clsDataSerializer` — Payload Builder

```vba
Function BuildPayload(boards, tasks) As String
    ' Returns: {"boards":[...],"tasks":[...]}
End Function
```

---

### Type Conversion

#### `mod_TypeConverter` — Safe Conversions

```vba
Function SafeLong(v As Variant) As Long          ' → 0 on error/null
Function SafeInt(v As Variant) As Integer        ' → 0 on error/null
Function SafeString(v As Variant) As String      ' → "" on null
Function FormatDateISO(d As Variant) As String   ' → "yyyy-mm-dd"
Function FormatTimeHM(t As Variant) As String    ' → "hh:mm"
```

Used everywhere to prevent type errors and null crashes.

---

### HTML Bridge

#### `clsHtmlBridge` — Browser Control Interface

Abstraction over the WebBrowser/Edge control for JS ↔ VBA communication.

```vba
Public Sub Initialize(wb As Object)
    ' Store reference to form's control
End Sub

Public Sub ExecuteJS(code As String)
    ' Run JavaScript: e.g., ExecuteJS "window.loadData(...)"
End Sub

Public Function RetrieveValue(jsExpr As String) As String
    ' Evaluate JS expression and get result: e.g., RetrieveValue("getPendingCommand()")
End Function

Public Function GetPendingCommand() As String
    ' Poll for pending JSON command from JavaScript
End Function

Public Sub PushData(jsonPayload As String, Optional restoreBoardId As String)
    ' Load data into page: window.loadData(payload); window.switchBoard(id);
End Sub

Public Function IsDocumentReady(url As Variant) As Boolean
    ' Check if document is our index.html
End Function
```

---

## Frontend Architecture

### HTML File: `kanban.html`

Contains the complete Kanban board UI with embedded CSS and ES6 JavaScript.

**Structure:**

```html

<header>           <!-- Board tabs, theme toggle, settings -->
<main id="main">   <!-- Dynamic column rendering -->
<modal id="taskModal">  <!-- Task editor modal -->
<script src="shared.js">
<script>           <!-- ES6 classes: KanbanApp, TaskEditor, etc. -->
```

---

### Styling: Embedded CSS

**Root Variables (Light/Dark Theme):**

```css

:root {
  --bg: #0f1117;          /* Background */
  --surface: #1a1d2e;     /* Cards, panels */
  --card: #252840;        /* Task cards */
  --border: #2e3250;      /* Borders */
  --accent: #6c63ff;      /* Primary color */
  --text: #e2e8f0;        /* Text color */
  --muted: #8892a4;       /* Muted text */
  --done: #10b981;        /* Success color */
}
```

**Responsive Design:**

- Flexbox columns (1 column per status)
- Horizontal scroll on small screens
- Modal dialogs overlay with fade
- Drag-drop visual feedback

---

### JavaScript Architecture

#### Class: `KanbanApp` (Main Entry Point)

```javascript
class KanbanApp {
  constructor() {
    this.boards = [];
    this.tasks = [];
    this.activeBoardId = null;
    this.bridge = new ParentBridge();     // VBA communication
    
    this.editor = new TaskEditor(this);   // Modal editor
    this.drag = new DragDropManager(this); // Drag-drop
    this.renderer = new BoardRenderer(this);
  }

  persist(action, payload) {
    // Send command to VBA
    this.bridge.persist(action, payload);
  }
}

const kanbanApp = new KanbanApp();
```

**Global Window API (exposed for onclick handlers):**

```javascript

window.loadData = (json) => { }           // Load from VBA
window.switchBoard = (id) => { }          // Change active board
window.openAddTask = (status) => { }      // Open add modal
window.openEditTask = (id) => { }         // Open edit modal
window.saveTask = () => { }               // Save & persist
window.deleteTask = (id) => { }           // Delete & persist
window.onDragStart/End/Over/Leave/Drop    // Drag-drop handlers
```

---

#### Class: `TaskEditor` — Modal Dialog

Manages task creation/editing in a wide modal.

```javascript
openAdd(status)       // New task, set default status
openEdit(id)          // Edit existing task
save()                // Collect form, call persist()
getSubtasks(parentId) // Filter child tasks
renderSubtaskList()   // Display subtask editor section
saveSubtask(id)       // Save child task
deleteSubtask(id)     // Delete child task
computeParentPct()    // Auto-calculate % from children
updateParentPct()     // Update read-only % input
```

##### Key Feature: Parent Percentage

- If task has subtasks: % is read-only, auto-calculated from children
- If no subtasks: % is editable

---

#### Class: `DragDropManager` — Drag & Drop

Implements HTML5 drag-drop for task reordering across columns.

```javascript
onStart(e, id)   // Store drag ID, add .dragging class
onDrop(e, colId) // Update task.status & sortOrder
onOver/onLeave   // Visual feedback (.drag-over class)
```

##### Feature: Automatic Reordering

- Task moved to new column
- `sortOrder` set to max(sortOrder in column) + 1
- `persist('updateTask', task)` sent to VBA

---

#### Class: `BoardRenderer` — Render Engine

Renders 5 Kanban columns from in-memory task array.

```javascript
render()         // Re-render all columns
card(task)       // Render individual task card
```

**Card Features:**

- Title, description, metadata badges
- Priority color coding
- Task type label
- Date range with time
- Subtask count
- Progress bar (if task or parent)
- Edit/delete buttons
- Overdue visual indicator
- Custom color left border

---

#### Class: `ParentBridge` — VBA Communication

Bridges JavaScript to VBA parent window.

```javascript
persist(action, payload) {
  // Call parent's setPendingCommand(json)
  if (this.inIframe && window.parent.setPendingCommand) {
    window.parent.setPendingCommand(JSON.stringify({action, payload}));
  }
}

setStatus(msg, isErr) {
  // Update parent window status bar
}
```

---

### Shared Utilities: `shared.js`

**HTML Escaping:**

```javascript
esc(s) → document.createElement('div').innerHTML = s
```

Prevents XSS when rendering user input (titles, descriptions).

**Date Helpers:**

```javascript
fmtDate(d)        → "yyyy-mm-dd"
parseDate(s)      → Date object
addDays(d, n)     → New date + n days
daysBetween(a, b) → Number of days
```

**Color Helpers:**

```javascript
lightenColor(hex, ratio)  → Lighter shade
contrastColor(hex)        → Black or white for readability
```

**Modal Helpers:**

```javascript
openModal(id)   → Add .open class
closeModal(id)  → Remove .open class
```

---

## Communication Bridge

### Flow: JavaScript → VBA

#### 1. User Action (e.g., Save Task)

```javascript
// In TaskEditor.save()
this.app.persist('addTask', {title: "...", ...});
```

#### 2. ParentBridge Sends Command

```javascript
// ParentBridge.persist()
if (window.parent && window.parent.setPendingCommand) {
  window.parent.setPendingCommand(JSON.stringify({
    action: 'addTask',
    payload: { title: "...", boardId: 1, ... }
  }));
}
```

#### 3. VBA Form Timer Polls

```vba
' In form's Timer event (every 200ms)
Dim cmd As String
cmd = m_Bridge.GetPendingCommand()

If cmd <> "" Then
  Dim result As String
  result = m_Router.Route(cmd)
  
  If result <> "NORELOAD" Then
    ' Reload and restore board context
    LoadAndRender result
  End If
  
  m_Bridge.ClearPendingCommand()
End If
```

#### 4. Command Router Dispatches

```vba
' clsCommandRouter.Route()
Select Case action
  Case "addTask"
    Route = TaskService.AddTask(jsonStr)
    ' Service handles DB insert
End Select
```

#### 5. VBA Pushes Updated Data Back

```vba
' Load all data and send to HTML
Dim payload As String
payload = m_Serializer.BuildPayload(allBoards, allTasks)

' Execute JS to load data
m_Bridge.PushData payload, result
```

#### 6. JavaScript Receives Data

```javascript
// In kanban.html
window.loadData = (json) => {
  const data = typeof json === 'string' ? JSON.parse(json) : json;
  app.boards = data.boards;
  app.tasks = data.tasks;
  app.renderer.render();  // Re-render UI
};
```

---

### Flow: VBA → JavaScript (Explicit Push)

When VBA needs to update UI without waiting for user action:

```vba
m_Bridge.ExecuteJS "window.setTheme('light');"
m_Bridge.ExecuteJS "window.switchBoard('2');"
```

---

## Data Flow Examples

### Example 1: Add New Task

```text

User clicks "Add Task" in Backlog column
  ↓
openAddTask('backlog') called
  ↓
TaskEditor.openAdd('backlog') opens modal
  ↓
User fills form, clicks "Save Task"
  ↓
TaskEditor.save() collects form data
  ↓
this.app.persist('addTask', {
  title: "Implement login",
  status: "backlog",
  boardId: 1,
  priority: "High",
  pctComplete: 0,
  ...
})
  ↓
ParentBridge.persist('addTask', payload)
  ↓
window.parent.setPendingCommand(json)
  ↓
VBA form timer polls getPendingCommand()
  ↓
clsCommandRouter.Route(json)
  ↓
clsTaskService.AddTask(json)
  ↓
clsTask.LoadFromJson(json)
clsTaskRepo.Add(task)
  ↓
INSERT INTO tblTasks (...)
  ↓
Return new activeBoardId
  ↓
VBA: SELECT all boards & tasks
  ↓
clsDataSerializer.BuildPayload(boards, tasks) → JSON
  ↓
clsHtmlBridge.PushData(json, boardId)
  ↓
ExecuteJS: window.loadData(json); window.switchBoard(boardId);
  ↓
JavaScript:
  app.boards = [...];
  app.tasks = [...];
  app.renderer.render();
  ↓
UI updates: New task appears in Backlog column
```

---

### Example 2: Drag Task to Another Column

```text
User drags task from "To Do" to "In Progress"
  ↓
onDragStart(e, taskId) called
  ↓
onDrop(e, 'inprog') called
  ↓
DragDropManager.onDrop():
  - Find task by ID
  - Set task.status = 'inprog'
  - Calculate new sortOrder (max existing + 1)
  ↓
this.app.persist('updateTask', task)
  ↓
[Same JSON→VBA→DB→JSON→JS flow as above]
  ↓
clsTaskService.UpdateTask(json)
  ↓
UPDATE tblTasks SET Status='inprog', SortOrder=5 WHERE TaskID=42
  ↓
Full data reload sent back to JS
  ↓
UI re-renders: Task appears in "In Progress" column
```

---

### Example 3: Add Subtask

```text

User editing task, sees "Subtasks" section on right
User clicks "+ Add Subtask"
  ↓
addSubtaskRow() creates temporary UI row
  ↓
User fills subtask form, clicks "Save Subtask"
  ↓
saveSubtask(tempId):
  - Collect subtask data from form
  - Set parentId = editingTaskId
  - Create new object: { id: 't' + timestamp, parentId, ... }
  - this.app.tasks.push(newSubtask)
  - this.app.persist('addTask', newSubtask)
  ↓
[JSON→VBA flow]
  ↓
clsTask.LoadFromJson sets parentId field
  ↓
INSERT INTO tblTasks (ParentID=original_taskid, ...)
  ↓
Data reloaded
  ↓
updateParentPct(editingTaskId):
  - Recalculate parent % from all subtasks
  - Set parent % input to read-only with computed value
  ↓
UI re-renders: Subtask appears in list, parent % updates
```

---

## Key Features

### 1. Kanban Board Workflow

**5 Columns (Statuses):**

- **Backlog** (#64748b gray) — Not started
- **To Do** (#3b82f6 blue) — Planned
- **In Progress** (#f59e0b amber) — Being worked on
- **Review** (#8b5cf6 purple) — Under review
- **Done** (#10b981 green) — Completed

**Interaction:**

- Drag tasks between columns to change status
- Tasks automatically reorder within column
- Task count shown in column header

---

### 2. Task Editor Modal

**Fields:**

- Title (required)
- Description (multiline)
- Notes (internal, multiline)
- Priority (Low/Medium/High) — Color-coded badge
- Status (dropdown, default based on which column's "+" clicked)
- Task Type (e.g., "Bug", "Feature", "Story")
- Start/End dates
- Start/End times
- % Complete (read-only if parent with subtasks)
- Card Color (custom hex or clear to default)

**Subtasks (for non-child tasks):**

- Right panel shows all child tasks
- Each subtask is collapsible editor
- Parent % auto-calculated from children
- Cascading deletes

---

### 3. Subtask Management

**Hierarchy:**

- Task can have many children (subtasks)
- Each subtask is a full task with all fields
- Only top-level tasks render in columns
- Subtasks only visible in parent's edit modal

**Parent % Complete:**

- Auto-calculated: Σ(child.pctComplete) / count
- Read-only in parent editor
- Subtasks can have independent progress

**Progress Bar:**

```text
[████████░░░░░░░░░] 45%
```

Rendered on both task cards and subtask rows.

---

### 4. Date & Time Features

**Fields:**

- Start Date / End Date (ISO format: YYYY-MM-DD)
- Start Time / End Time (HH:MM format)

**Display:**

- Card shows date range: "Apr 8 – Apr 15"
- If time set: "09:00 Apr 8 – Apr 15"
- **Overdue Detection:** If endDate < today and status ≠ done, card gets red border & highlight

**Badge Color:**

- Normal dates: gray
- Overdue dates: red (#ef4444) background

---

### 5. Priority & Type Badges

**Priority (3 levels):**

- Low (#10b981 green)
- Medium (#f59e0b amber) — default
- High (#ef4444 red)

**Type Badge:**

- Custom text label (e.g., "Bug", "Feature")
- Purple background, displayed on card

**Subtask Count:**

- If parent task: 📎 N badge showing subtask count

---

### 6. Custom Card Colors

**Feature:**

- Color picker in modal
- 3px left border on task card
- Can be cleared (empty string) to remove color
- Independent of priority color

---

### 7. Drag & Drop Reordering

**Mechanism:**

- HTML5 Drag API
- Task dragged within column or to another column
- Drop creates new SortOrder

**Visual Feedback:**

- Source card: .dragging class (opacity: 0.35)
- Drop zone: .drag-over class (blue dashed border)

**Backend:**

- Creates updateTask command
- New SortOrder = max(sortOrder in destination column) + 1

---

### 8. Theme Switching

**Two Themes:**

- Dark (default)
- Light

**Mechanism:**

- HTML attribute: `<html data-theme="dark">`
- CSS variables change based on `[data-theme]` selector
- Button in header toggles theme
- Calls VBA `setTheme()` function for persistence

---

### 9. Board Management

**Create Board:**

- Header button: "+ Board"
- Modal form input board name & color
- Persists as new row in tblBoards

**Board Tabs:**

- Each board is a tab in header
- Active tab highlighted with accent color
- Delete button (×) on each tab (soft-delete cascade)

**Switch Boards:**

- Click tab or programmatically `window.switchBoard(boardId)`
- Filters tasks to only this boardId
- State: `app.activeBoardId`

---

### 10. Data Persistence

**Architecture:**

- In-memory state in JavaScript (app.boards, app.tasks)
- Every change triggers persist() → VBA
- VBA updates database
- Full data payload sent back
- JavaScript state refreshed from DB payload

**No Local Storage:**
Data lives only in Access database (tblBoards, tblTasks)

---

## Database Design (Quick Reference)

### tblBoards

```VBA
BoardID (PK)        Long
BoardName           Text
BoardColor          Text (hex color)
CreatedOn           DateTime
ModifiedOn          DateTime
IsDeleted           Boolean
```

### tblTasks

```VBA
TaskID (PK)         Long
BoardID (FK)        Long → tblBoards
Title               Text (required)
Description         Text
Priority            Text
Status              Text
StartDate           DateTime
EndDate             DateTime
StartTime           DateTime
EndTime             DateTime
PctComplete         Integer (0-100)
TaskColor           Text (hex or null)
TaskType            Text
ParentID (FK)       Long → tblTasks (nullable for subtasks)
Notes               Text
IsExpanded          Boolean (collapse subtasks)
SortOrder           Long (position in column)
CreatedOn           DateTime
ModifiedOn          DateTime
IsDeleted           Boolean
```

---

## Summary

**This architecture achieves:**

✅ **Separation of Concerns:** VBA handles persistence, JS handles UI  
✅ **JSON-Based Communication:** Language-agnostic, human-readable  
✅ **Responsive UI:** Instant feedback, no page reloads  
✅ **Rich Features:** Drag-drop, subtasks, date tracking, priority levels  
✅ **Type Safety:** Safe conversions prevent crashes (mod_TypeConverter)  
✅ **Clean Code:** Classes use Service + Repository patterns  
✅ **Extensibility:** Adding new fields is straightforward (add column, update DTO classes)

The frontend polls VBA for commands, executes them, and pulls fresh data back—enabling smooth interactivity with full database integrity.
