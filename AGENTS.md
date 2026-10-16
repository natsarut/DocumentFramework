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
- **Tooling:** template generation, scaffolding, and developer productivity tools

## Solution Map
- `TemplateGenerator`
- `DocumentOrchestratorService`
- `TemplateModelService`
- `TemplateEngineService`
- `DocumentConverterService`