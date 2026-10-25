# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Project

Early-stage JavaScript project. The only dependency so far is
[Lenis](https://github.com/darkroomengineering/lenis) (smooth scrolling).

## Commands

- `npm install` — install dependencies (runs automatically at the start of
  Claude Code on the web sessions via `.claude/hooks/session-start.sh`)

No build, test, or lint scripts are defined yet; add them to `package.json`
and list them here as they're introduced.

## Conventions

- `node_modules/` is git-ignored; commit `package-lock.json` changes alongside
  `package.json` changes.
