# Vision

This project provides a Model Context Protocol (MCP) service for Claude that allows you to interact
with your Microsoft Todo tasks using natural language.

# Microsoft Todo Hierarchy

Microsoft Todo is organized in a three-level hierarchy:

1. **Task Lists** - The top-level containers that organize tasks into categories, projects, or areas
   of focus (e.g., "Work", "Personal", "Shopping"). These help you group related tasks together.

2. **Tasks** - The main todo items that represent actions or activities you need to complete. Tasks
   have properties like title, description, due date, importance, and status.

3. **Checklist Items** - Subtasks that belong to a parent task. These allow you to break down a task
   into smaller, manageable steps or components.

# Supported Actions

The current implementation supports the following actions:

## Task List Management

1. `get-task-lists` - Get all Microsoft Todo task lists (top-level containers)
   - Shows list name, ID, and additional information (default list, shared status)
   - Use this to find the IDs needed for other commands

2. `create-task-list` - Create a new task list to organize your tasks
   - Required: displayName
   - Example: `@mstodo create-task-list displayName="Work Projects"`

3. `update-task-list` - Rename an existing task list
   - Required: listId, displayName
   - Example: `@mstodo update-task-list listId="LIST_ID" displayName="Important Work Projects"`

4. `delete-task-list` - Delete a task list and all tasks within it
   - Required: listId
   - Example: `@mstodo delete-task-list listId="LIST_ID"`

## Task Management

5. `get-tasks` - Get tasks from a specific list
   - Required: listId
   - Optional: filter, select, orderby, top, skip, count
   - Supports OData query parameters for advanced filtering and sorting
   - Example: `@mstodo get-tasks listId="LIST_ID" filter="status eq 'notStarted'"`

6. `create-task` - Create a new task in a specific list
   - Required: listId, title
   - Optional: body, dueDateTime, startDateTime, importance, isReminderOn, reminderDateTime, status,
     categories
   - Example:
     `@mstodo create-task listId="LIST_ID" title="Finish report" dueDateTime="2023-12-31T23:59:59Z"`

7. `update-task` - Update an existing task
   - Required: listId, taskId
   - Optional: title, body, dueDateTime, startDateTime, importance, isReminderOn, reminderDateTime,
     status, categories
   - Note: Empty string for date fields will remove the date
   - Example: `@mstodo update-task listId="LIST_ID" taskId="TASK_ID" status="completed"`

8. `delete-task` - Delete a task and all its checklist items
   - Required: listId, taskId
   - Example: `@mstodo delete-task listId="LIST_ID" taskId="TASK_ID"`

## Checklist Item Management (Subtasks)

9. `get-checklist-items` - Get subtasks for a specific task
   - Required: listId, taskId
   - Displays: item name, status (completed/not completed), creation date, and ID
   - Shows the parent task title for better context
   - Example: `@mstodo get-checklist-items listId="LIST_ID" taskId="TASK_ID"`

10. `create-checklist-item` - Create a new subtask for a task
    - Required: listId, taskId, displayName
    - Optional: isChecked (default is false)
    - Example:
      `@mstodo create-checklist-item listId="LIST_ID" taskId="TASK_ID" displayName="Research competitors"`

11. `update-checklist-item` - Update an existing subtask
    - Required: listId, taskId, checklistItemId
    - Optional: displayName, isChecked (at least one required)
    - Example:
      `@mstodo update-checklist-item listId="LIST_ID" taskId="TASK_ID" checklistItemId="ITEM_ID" isChecked=true`

12. `delete-checklist-item` - Delete a subtask from a task
    - Required: listId, taskId, checklistItemId
    - Example:
      `@mstodo delete-checklist-item listId="LIST_ID" taskId="TASK_ID" checklistItemId="ITEM_ID"`

## Authentication

13. `auth-status` - Check if you're authenticated with Microsoft Graph API
    - Shows token status and expiration time
    - Example: `@mstodo auth-status`

# references

- [MCP Documentation](mcp.md)
- [Microsoft TODOs](microsofttodo.md)
  - [todotasklist](todotasklist.md)
  - [todotask](todotask.md)

# memo

https://smithery.ai/?q=todo https://glama.ai/mcp/servers?query=todo&sort=search-relevance%3Adesc

Higher-Level Tools We Could Implement: create-smart-task - Create a main task with AI-generated
subtasks breakdown-task - Take an existing task and break it down into subtasks add-today-task -
Quickly add a task to your "Today" list without needing to know list IDs complete-task - Mark a task
as complete (simpler than using update-task) list-my-tasks - Show tasks across lists with filtering
by status/date without requiring complex OData queries reschedule-task - Push a task's due date
forward by a certain number of days prioritize-tasks - Automatically sort tasks by importance, due
date, etc.

Options to Make It Compatible To make this work more like standard MCP servers, you'd need to:
Modify the authentication approach: Package the auth logic directly into the MCP server Add a way to
handle first-time auth within the server startup Support environment variable-based token injection
like the GitHub example Package it properly: Create a proper npm package that includes both the
server and auth logic Set up the package.json to make it runnable via npx For now, the best approach
is to follow the current setup steps - install locally, run the separate auth process, and then
configure Claude to point to the local build.
