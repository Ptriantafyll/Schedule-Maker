# CLAUDE.md

## Project Overview

Schedule-Maker is a Python tool that generates optimized monthly on-duty schedules for hospital doctors. It uses Google OR-Tools CP-SAT constraint programming solver to produce fair, balanced duty assignments while respecting doctor unavailability and quality-of-life preferences.

## Preferences

- When making changes don't paste all the code at once. Instead go step by step in small chunks of code explaining the process each time
- When the user is learning and doing most of the coding themselves, explain: (1) why we are making a decision, (2) what the best practices are and (3) if there is a new feature that we haven't touched explain how it works
- When the user stops and corrects a suggestion or says they don't like something, add that preference to this CLAUDE.md file
- After every change, check if anything can be made cleaner and if there is repeated code that can be extracted into a helper function
- Use snake_case for OR-Tools CP-SAT methods (e.g. `model.add`, `model.new_bool_var`, `only_enforce_if`), not PascalCase
- Use TDD (Test-Driven Development) approach: write tests first, then implement the code to make them pass
