## Purpose
This file helps AI coding agents work effectively in this repository. Use it as a fast orientation guide before making changes.

## Project Overview
This repository provides a production-grade document solution designed in strict accordance with Microservices, Clean Architecture using repository pattern principles and modern enterprise application standards.

Built on .NET 10, the solution demonstrates a well-structured, scalable, and maintainable architecture for developing complex business systems. It integrates advanced document generation capabilities, AI-assisted development workflows, and specification-driven design patterns, enabling teams to accelerate development while preserving architectural consistency and code quality.

The solution is intended to serve both as a reference implementation for Clean Architecture best practices and as a ready-to-use foundation for enterprise-level applications that require long-term maintainability, extensibility, and high development efficiency.

The main application layers are:
- **Domain:** core entities, domain events, and business rules (no dependencies)
- **Application:** business logic, interfaces, use cases, DTOs, validation, pipeline behaviors, and application abstractions
- **Infrastructure:** External concerns (database, email, file system)
- **HTTP API & Web UI:** HTTP APIs, Blazor components and user interface

## Solution Map
- `src/DocumentFrameworkSolution/TemplateGenerator` Blazor Web App for converting Microsoft Word documents to HTML.
- `src/DocumentFrameworkSolution/TemplateGenerator/TemplateGenerator` Blazor Server for converting Microsoft Word documents to HTML.
- `src/DocumentFrameworkSolution/TemplateGenerator/TemplateGenerator.Client` Blazor Client for converting Microsoft Word documents to HTML.

## Working Guidance
- Preserve Clean Architecture boundaries. Avoid introducing UI or infrastructure concerns into `Domain`.
- Prefer existing repository patterns over inventing new ones. Match nearby features before creating new abstractions.
- When adding a new feature, inspect similar modules in `Application`, `Infrastructure`, and `Server.UI` first.
- Keep changes targeted. Do not refactor unrelated areas unless the task requires it.
- Update documentation when setup steps, commands, or workflows change.
- If a task benefits from explicit design or planning, add artifacts under `docs/superpowers/` rather than inventing a separate process.

## Common Commands
Build the solution: `dotnet build DocumentFrameworkSolution.slnx`

Run tests: `dotnet test DocumentFrameworkSolution.slnx`

## Where To Look First
- `README.md` for project overview, setup, and supported workflows
- `docs/` for detailed design, architecture, and development guidelines
- `docs/superpowers/` for prior design notes and implementation plans when available

## Expected Output Style
- Use clear, concise, and professional language.
- Favor small, reviewable edits.
- Explain assumptions when behavior is unclear.
- Verify builds or tests when the task meaningfully changes behavior.
- Keep new guidance aligned with the current repository structure instead of generic best practices
- Use code snippets to illustrate examples when appropriate.
- When suggesting changes, provide reasoning and context to help reviewers understand the motivation behind the edits.
- When proposing new features or significant changes, include a brief design rationale and potential impact on existing architecture.
