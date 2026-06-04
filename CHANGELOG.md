# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.25.1] - 2026-06-04

### Changed
- Documentation only: added npm badges (version, downloads, node, license) and pointed all clone, install and repository references at `@shayanline/msteams-mcp`.

## [0.25.0] - 2026-06-04

### Added
- Messaging: `importance` (high/urgent) and `subject` on `teams_send_message`, plus `teams_forward_message`.
- Browse: `teams_list_chats`, `teams_list_teams`, `teams_mark_unread`, `teams_get_chat_members`.
- Chat management: `teams_rename_chat`, `teams_add_member`, `teams_remove_member`, `teams_leave_chat`, `teams_pin_message`, `teams_unpin_message`, `teams_mute_chat`, `teams_unmute_chat`.
- Scheduling: `teams_find_meeting_times`.
- Files: `teams_list_files`, `teams_upload_file`, `teams_download_file`, `teams_send_file`.
- Microsoft To Do: `teams_list_task_lists`, `teams_list_tasks`, `teams_create_task`, `teams_complete_task`.

### Changed
- Project is now maintained and released independently as `@shayanline/msteams-mcp`. `main` is the source of truth.

## [0.24.1] and earlier

See the [GitHub releases](https://github.com/shayanline/msteams-mcp/releases) for the history before independent maintenance began.

[Unreleased]: https://github.com/shayanline/msteams-mcp/compare/v0.25.1...HEAD
[0.25.1]: https://github.com/shayanline/msteams-mcp/compare/v0.25.0...v0.25.1
[0.25.0]: https://github.com/shayanline/msteams-mcp/compare/v0.24.1...v0.25.0
