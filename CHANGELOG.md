# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.28.0] - 2026-06-18

### Added
- `teams_send_files`: send several local files into a conversation as native attachments on a **single message** (one post, multiple file chiclets), with an optional caption. Use it instead of calling `teams_send_file` once per file when the files should be grouped under one message. Files are uploaded to the right place per conversation type, the same as `teams_send_file`.

## [0.27.0] - 2026-06-16

### Added
- `teams_get_recording`: get a meeting's recording(s) from the meeting `threadId` (from `teams_get_meetings`). Returns metadata (title, file name, start/end time, duration) plus a `playbackUrl` (opens in the SharePoint/Stream web player) and a `downloadUrl` (direct `.mp4`) for each recording, rather than the video bytes. Uses the same Substrate `WorkingSetFiles` lookup as `teams_get_transcript`, so no extra authentication is needed.

## [0.26.1] - 2026-06-15

### Fixed
- Markdown to Teams HTML: a list, blockquote or table now renders correctly when it immediately follows a text line within the same block (e.g. a `Label:` line directly above `- bullets`), without needing a separating blank line. The converter previously required every line in a block to share one type, so a label plus a list fell back to a single `<br>`-joined paragraph and the bullets did not render. Blocks are now parsed by grouping consecutive line runs, letting block elements interrupt a paragraph as standard markdown allows.

## [0.26.0] - 2026-06-09

### Added
- `teams_send_file` now sends a **native file attachment** (a real file chiclet that also shows in the conversation's Files tab) instead of posting a share link. The file is uploaded to the right place per conversation type, matching the Teams client: a channel's own SharePoint files folder for channels, or OneDrive "Microsoft Teams Chat Files" (shared via an organisation link) for 1:1, group, meeting and self chats.
- `teams_create_channel` and `teams_delete_channel` to create and delete team channels (standard or private) via the middle-tier API.

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

[Unreleased]: https://github.com/shayanline/msteams-mcp/compare/v0.28.0...HEAD
[0.28.0]: https://github.com/shayanline/msteams-mcp/compare/v0.27.0...v0.28.0
[0.27.0]: https://github.com/shayanline/msteams-mcp/compare/v0.26.1...v0.27.0
[0.26.1]: https://github.com/shayanline/msteams-mcp/compare/v0.26.0...v0.26.1
[0.26.0]: https://github.com/shayanline/msteams-mcp/compare/v0.25.1...v0.26.0
[0.25.1]: https://github.com/shayanline/msteams-mcp/compare/v0.25.0...v0.25.1
[0.25.0]: https://github.com/shayanline/msteams-mcp/compare/v0.24.1...v0.25.0
