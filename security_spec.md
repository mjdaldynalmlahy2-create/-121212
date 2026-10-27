# Security Specification for Integrated HR & Performance System

## Data Invariants
1. A user profile must exist for every authenticated user to access role-protected data.
2. The first user to register in an empty system is granted the `admin` role.
3. Employee records are the master source of truth for identity in the system.
4. Tasks can only be updated in specific ways depending on the user's role (Assignee vs. Supervisor/Admin).

## The "Dirty Dozen" Payloads
1. **Self-Promotion**: An `employee` tries to update their own role to `admin` in `users` collection.
2. **Shadow Task Update**: An assignee tries to update the `title` or `priority` of a task they are assigned to.
3. **Orphaned Profile**: A user tries to create a user profile for a non-existent `employeeId`.
4. **ID Poisoning**: A user tries to create a custom ID for a task that is a 1MB string.
5. **PII Leak**: A non-admin user tries to `list` all users to scrape phone numbers and emails.
6. **Clock-In Spoof**: A user tries to create an attendance record for another user.
7. **Accrual Manipulation**: An employee tries to update their own accrual balance.
8. **Evaluation Bypass**: An employee tries to create or update their own evaluation.
9. **Status Shortcutting**: An employee transitions a task from `pending` straight to `fully_approved`.
10. **Admin Locked Out**: A user tries to delete the last admin (though rules might allow delete, app logic should prevent).
11. **System Config Tamper**: A non-admin tries to update `companyInfo`.
12. **Ghost Task Creation**: A non-authorized role tries to create a task.

## Red Team Audit Results
- **Identity Spoofing**: Prevented by `userId == request.auth.uid` and `isAdmin()` gates.
- **State Shortcutting**: Controlled by `affectedKeys().hasOnly()` on task updates.
- **Resource Poisoning**: `isValidId()` check on paths.
- **PII Isolation**: `users` list restricted to `isAdmin() || isGM()`.

## Test Runner (Conceptual)
Verify that all "Dirty Dozen" payloads return `PERMISSION_DENIED`.
