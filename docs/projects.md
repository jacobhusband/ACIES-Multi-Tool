# Projects Tab Guide

This tab is the project dashboard: it lists projects, lets you filter and search, and provides entry points to create, edit, review, and export data.

## Layout
- **Toolbar**: Search bar, help, quick new project, AI project creation, and stats shortcut.
- **Filters**: Timeframe chips (All, Past Weeks, This Week, Upcoming Weeks) and Status chips (All, Incomplete, Waiting, Working, Pending Review, Complete, Delivered).
- **Table**: ID, project details, due date, status, tasks/notes, and actions.
- **Empty state**: Prompts you to create the first project.

## Common Actions
- **Search**: Type in the search box to filter projects by text.
- **Create**: Click “+” for a quick blank project or the robot icon to generate from email/AI.
- **Sort**: Click column headers (ID, Project Details, Due) to sort; sorting respects current filters.
- **Filter**: Click any chip; the active chip stays highlighted.
- **Edit**: Use the row actions to edit project info, tasks, links, and status tags.
- **Status tags**: Apply the predefined statuses; they also drive filters and sorting priority.
- **Tasks/notes**: Add two links per task; links are normalized to local paths or URLs.
- **Open path**: Where available, the action opens the project folder via the backend helper.
- **Statistics**: The chart icon opens aggregate counts and averages.

## Tips
- Date parsing accepts `YYYY-MM-DD`, `MM/DD/YY`, `MM/DD/YYYY`, and tries natural Date parsing.
- “Past Weeks/This Week/Upcoming Weeks” are relative to today; weekly buckets reset each Sunday.
- Overdue items get a pill indicator; use “Mark overdue” tools to bulk update.
- Data is persisted via the backend; edits auto-save on confirm.
