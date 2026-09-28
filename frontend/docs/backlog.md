# Frontend Backlog

Deferred issues that should be addressed outside the current implementation step.

## Open

### BL-001: Self-Service Password Reset Flow (Full-Stack)

**Area:** Authentication / Self-Service Account Recovery  
**Priority:** Medium

#### Problem

Currently, hospital doctors and staff who forget their login credentials must contact an administrator or the IT Help Desk (Ext. 4420) to have their credentials reset manually. While the Login screen displays a Help Desk modal upon tapping *"Forgot Password?"*, an automated, self-service recovery mechanism is necessary for production scaling and operational convenience across hospital shifts.

Because password reset requires cryptographically secure, time-limited single-use reset tokens and automated email delivery, implementing this capability requires coordinated full-stack changes across both the FastAPI backend and the Flutter frontend.

#### Required work

1. **Backend Implementation**:
   - Create reset token model or signed cryptographic token generator (e.g. `itsdangerous` or hashed database tokens with strict expiration, e.g. 15–30 minutes).
   - Create request endpoint: `POST /api/v1/auth/forgot-password` accepting user email. Always return a generic 200/202 success message to prevent user enumeration attacks.
   - Integrate with email dispatch service (see backend backlog `BL-002`) to email the reset link containing the token.
   - Create confirmation endpoint: `POST /api/v1/auth/reset-password` accepting token and new password. Validates token validity/expiry, updates password hash in database, invalidates existing user refresh tokens/sessions, and revokes the used token.
   - Pytest unit and integration test suite covering token generation, valid reset, expired token, reused token, and invalid email.

2. **Frontend Implementation**:
   - **Data Layer**:
     - Add `forgotPassword(String email)` and `resetPassword({required String token, required String newPassword})` to `AuthRemoteDataSource` and `AuthRepository`.
   - **Presentation Layer**:
     - Connect *"Forgot Password?"* on `LoginScreen` to open a "Reset Password" dialog or dedicated screen.
     - Email request step: input email with validation, submitting to `forgotPassword`.
     - Token & New Password step: form input with password visibility toggle and confirmation password validation.
     - Provide clear feedback to the user on completion and route back to `LoginScreen`.
   - **Widget and Controller Tests**:
     - Comprehensive unit and widget tests covering input validation, submission states, and error handling.

#### Completion criteria

- Staff members can independently request a password reset via email.
- Password reset tokens are single-use, cryptographically secure, and expire after 15–30 minutes.
- User enumeration is prevented by returning generic responses on reset requests.
- Passwords can be successfully reset, allowing login with the new credentials.
- All backend and frontend test suites pass with 100% verification.
