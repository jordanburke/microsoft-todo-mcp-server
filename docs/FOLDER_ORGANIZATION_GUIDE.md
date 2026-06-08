# Microsoft To Do MCP Server - Folder Organization Enhancement

## Overview

While the Microsoft Graph API doesn't expose folder/group information for To Do lists, this MCP
server now includes enhanced tools that provide folder-like organization based on naming patterns
and list properties.

## New Tools

### 1. `get-task-lists-organized`

Get all task lists organized into logical folders/categories based on naming patterns, emoji
prefixes, and sharing status.

**Parameters:**

- `includeIds` (optional): Include list IDs in output (default: false)
- `groupBy` (optional): Grouping strategy - 'category' (default), 'shared', or 'type'

**Example Usage:**

```bash
# Get organized view
get-task-lists-organized

# Get organized view with IDs
get-task-lists-organized --includeIds true

# Group by sharing status
get-task-lists-organized --groupBy shared
```

**Organization Categories:**

- ⭐ **Special Lists**: Default task list and flagged emails
- 👥 **Shared Lists**: Lists shared with others
- 💼 **Work**: Lists starting with "Work" or "SBIR"
- 👪 **Family**: Lists with family emoji prefix
- 🏡 **Properties**: Lists with house emoji prefix
- 🛒 **Shopping Lists**: Lists with shopping cart emoji
- 🚗 **Travel & Rangeley**: Travel-related lists
- 🎉 **Seasonal & Events**: Holiday and event lists
- 📚 **Reading**: Reading lists
- 📦 **Archives**: Archive lists and lists with "(Location - Archived)" pattern
- 📋 **Other Lists**: Everything else

### 2. `archive-completed-tasks`

Move completed tasks older than a specified number of days from one list to another (archive) list.

**Parameters:**

- `sourceListId`: ID of the source list to archive tasks from
- `targetListId`: ID of the target archive list
- `olderThanDays`: Archive tasks completed more than this many days ago (default: 90)
- `dryRun`: If true, only preview what would be archived without making changes

**Example Usage:**

```bash
# Preview what would be archived
archive-completed-tasks --sourceListId "SOURCE_ID" --targetListId "TARGET_ID" --dryRun true

# Archive tasks older than 90 days
archive-completed-tasks --sourceListId "SOURCE_ID" --targetListId "TARGET_ID"

# Archive tasks older than 30 days
archive-completed-tasks --sourceListId "SOURCE_ID" --targetListId "TARGET_ID" --olderThanDays 30
```

## Naming Conventions for Organization

To take full advantage of the organized view, consider using these naming patterns:

### Emoji Prefixes

- 🛒 for shopping lists (Amazon, Grocery, Target)
- 🏡 for property-related lists
- 👪 for family lists
- 🎄🎉 for seasonal/event lists
- 📰 for reading lists
- 🚗 for travel lists
- 📦 for archive lists

### Archive Pattern

For archived lists from specific contexts, use: `Original Name (Context - Archived)`

Examples:

- "Home Hardware (Gore - Archived)"
- "Household Projects (Gore - Archived)"

## API Discovery Notes

During development, we discovered:

1. **Microsoft Graph To Do API (v1.0 and beta)** doesn't expose group/folder information
2. **Outlook Task Groups API** exists in beta but maps to the same flat list structure
3. All To Do lists belong to a single task group called "My Tasks"
4. The Microsoft To Do app's folder feature is a client-side organization not exposed via API

## Future Enhancements

If Microsoft adds folder support to the Graph API, consider:

1. Adding `get-task-groups` tool to retrieve native groups
2. Adding `create-task-group` tool to create new groups
3. Adding `move-list-to-group` tool to reorganize lists
4. Updating `get-task-lists-organized` to use native groups when available

## Technical Implementation

The organization logic uses:

- Regular expressions to detect naming patterns
- Priority-based sorting for consistent category display
- Hierarchical tree display with Unicode box drawing characters
- Smart categorization that checks multiple patterns in order

See `src/todo-index.ts` for the full implementation of the `organizeLists` function.
