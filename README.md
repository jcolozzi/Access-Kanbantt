# Access Kanbantt

Access Kanbantt is a Microsoft Access project management app that combines:

- Kanban board workflow
- Gantt timeline planning
- Analytics dashboard
- VBA + DAO backend services
- HTML/CSS/JavaScript frontend hosted in an Access Edge browser control

The app is designed for Access-first teams that want a modern planning UI while keeping data and business logic inside an Access database.

## What It Does

- Manage multiple boards and tasks with soft-delete semantics
- Move work across Kanban states (Backlog, To Do, In Progress, Review, Done)
- Plan and adjust date-based work in a Gantt timeline
- Track KPIs and schedule health in a dashboard view
- Support subtasks and parent progress rollup
- Push near-real-time updates across sessions using a change-log poller

## Tech Stack

- Host: Microsoft Access (.accdb)
- Data access: DAO (CurrentDb, Recordset)
- Backend architecture: DTO + Repository + Service + Router
- Frontend: HTML/CSS/JavaScript (ES6 classes)
- UI host: Edge browser control on form `frmKanbantt`
- Bridge: JavaScript pending-command queue polled by VBA timer

## Project Structure

```text
Access Kanbantt/
  Kanbantt.accdb
  data/
    Kanbantt_be.accdb
  view/
    index.html
    kanban.html
    gantt.html
    dashboard.html
    shared.js
  cls/
    frmKanbantt.frm
    frmPubSub.frm
    clsCommandRouter.cls
    clsHtmlBridge.cls
    clsDataSerializer.cls
    clsBoard.cls
    clsTask.cls
    clsBoardRepo.cls
    clsTaskRepo.cls
    clsBoardService.cls
    clsTaskService.cls
    clsPubSubBroker.cls
  bas/
    mod_JsonUtility.bas
    mod_TypeConverter.bas
    mod_ChangeLog.bas
```

## Runtime Architecture

### Core flow

1. `frmKanbantt` loads and navigates `WebBrowser1` to `view/index.html`.
2. JavaScript sends user actions as JSON commands via `setPendingCommand(...)`.
3. `Form_Timer` polls with `clsHtmlBridge.GetPendingCommand()`.
4. `clsCommandRouter.Route(...)` dispatches action to services.
5. Services call repositories to persist data in `tblBoards` / `tblTasks`.
6. `clsDataSerializer` rebuilds payload and `clsHtmlBridge.PushData(...)` refreshes UI.

### Important VBA classes

- `clsCommandRouter`: central action dispatcher
- `clsBoardService`, `clsTaskService`: domain logic
- `clsBoardRepo`, `clsTaskRepo`: DAO CRUD + soft delete + sort behavior
- `clsHtmlBridge`: JS execute/retrieve/push helpers
- `clsDataSerializer`: payload builder for boards/tasks JSON
- `clsPubSubBroker` + `frmPubSub`: poll and broadcast cross-session changes

### Frontend views

- `view/index.html`: shell, navigation, board tabs, settings, iframe host
- `view/kanban.html`: board columns, drag/drop, task modal, subtasks
- `view/gantt.html`: timeline chart, task table, date controls, zoom
- `view/dashboard.html`: KPI cards, status/priority charts, health views
- `view/shared.js`: utility helpers and parent bridge

## Data Model (High-Level)

- `tblBoards`: board identity, name, color, deletion flag
- `tblTasks`: task details, status, dates/times, progress, hierarchy, sort, deletion flag
- `tblChangeLog`: append-only change events used by the pub/sub poller

Both boards and tasks use soft-delete updates (`IsDeleted = True`) instead of hard deletes.

## Command Contract (JS -> VBA)

Commands handled by `clsCommandRouter` include:

- `setActiveBoard`
- `addBoard`, `renameBoard`, `updateBoard`, `deleteBoard`
- `addTask`, `updateTask`, `deleteTask`
- `toggleExpand`, `reorderTask`
- `updateAccent` (UI-only/no data reload)

The router returns either:

- a board id to reload/restore context, or
- `NORELOAD` when no full data refresh is needed

## Getting Started

1. Open `Kanbantt.accdb` in Microsoft Access on Windows.
2. Open form `frmKanbantt`.
3. Confirm the Edge control can navigate to `CurrentProject.Path\view\index.html`.
4. Create a board, then add tasks from Board or Gantt view.

## Multi-Session Sync

- `mod_ChangeLog.LogChange` records board/task add/update/delete actions.
- Hidden form `frmPubSub` polls `tblChangeLog` every 3 seconds.
- Changes from other users are emitted through broker events.
- `frmKanbantt` applies updates by executing JS merge functions.

By default, stale change-log entries are purged with:

- `PurgeChangeLog 7`

## Development Notes

- Keep backend action names aligned between JavaScript and `clsCommandRouter`.
- For new persistence features, follow this pattern:
  1. Add UI command emit in HTML/JS.
  2. Add router case.
  3. Add service method.
  4. Add repository method.
  5. Extend DTO serialization/deserialization if needed.
- Use `mod_TypeConverter` and `mod_JsonUtility` helpers to avoid null/type issues.

## Troubleshooting

- UI not loading:
  - verify `view/index.html` exists under the Access project path
  - verify `WebBrowser1_DocumentComplete` is firing
- Commands not persisting:
  - verify `Form_Timer` is enabled
  - verify `getPendingCommand()` / `clearPendingCommand()` JavaScript functions are available
- Cross-session updates not appearing:
  - verify hidden `frmPubSub` is open
  - verify `tblChangeLog` is receiving rows

## Related Docs

- `ARCHITECTURE.md`

