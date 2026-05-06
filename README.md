# Access Kanbantt

![Platform](https://img.shields.io/badge/Platform-Microsoft%20Access-00599C)
![UI](https://img.shields.io/badge/UI-Kanban%20%7C%20Gantt%20%7C%20Dashboard-6C63FF)
![Backend](https://img.shields.io/badge/Backend-VBA%20%2B%20DAO-1F6FEB)
![Sync](https://img.shields.io/badge/Sync-ChangeLog%20PubSub-10B981)
![Status](https://img.shields.io/badge/Status-Active-22C55E)

Access Kanbantt is a Microsoft Access project manager with a modern browser-based UI. It combines a Kanban board, Gantt timeline, and analytics dashboard while keeping data and business logic in Access.

## Why Teams Use It

- Plan work visually with board and timeline views
- Track progress, risk, and overdue work in one dashboard
- Keep existing Access workflows and data ownership
- Support multi-user updates with change-log based sync

## Core Views

- Board: drag/drop tasks across Backlog, To Do, In Progress, Review, Done
- Gantt: schedule and adjust task timelines
- Dashboard: KPI cards, status and priority breakdowns, board health

## Screenshots

![Kanban board screenshot placeholder](https://github.com/jcolozzi/Access-Kanbantt/blob/main/images/Kanban.png)
![Gantt timeline screenshot placeholder](https://github.com/jcolozzi/Access-Kanbantt/blob/main/images/Gantt.png)
![Dashboard screenshot placeholder](https://github.com/jcolozzi/Access-Kanbantt/blob/main/images/Dashboard.png)

## Quick Start

1. Open `Kanbantt.accdb` in Microsoft Access (Windows).
2. Open form `frmKanbantt`.
3. Create a board, then add tasks from Board or Gantt view.

## Repository Highlights

- `view/`: HTML/JS frontend (`index.html`, `kanban.html`, `gantt.html`, `dashboard.html`)
- `cls/`: VBA classes (router, services, repositories, bridge)
- `bas/`: shared JSON/type/change-log modules

## Documentation

- Technical and developer guide: [README.md](README.md)
