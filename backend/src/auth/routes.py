"""
Authentication routes
"""
import uuid
from typing import Optional

from fastapi import APIRouter, Depends, status, Response, Request
from fastapi.security import OAuth2PasswordRequestForm
from sqlmodel import Session

from src.auth import controllers as auth_controllers
from src.auth.schemas import Token
from src.auth.dependencies import (
    get_current_user,
    require_admin,
    require_department_admin,
    require_super_admin,
)
from src.user.models import User as UserModel
from src.user.schemas import UserRead
from src.db.connection import get_session

from src.auth.schemas import (
    InvitationRead,
    InvitationCreatedResponse,
    DepartmentProvisioningResponse,
    StaffInvitationCreate,
    DepartmentAdminProvisioningCreate,
    DepartmentAdminInvitationCreate,
    InvitationSignupRequest,
    RefreshTokenRequest,
)

router = APIRouter(
    prefix="/auth",
    tags=["Authentication"]
)


@router.get("/me", response_model=UserRead, status_code=status.HTTP_200_OK)
def get_current_user_profile(current_user: UserModel = Depends(get_current_user)):
    """Returns the profile of the currently authenticated user."""
    return current_user


@router.post("/login", response_model=Token, status_code=status.HTTP_200_OK)
def login(
    form_data: OAuth2PasswordRequestForm = Depends(),
    session: Session = Depends(get_session),
    response: Response = None,
):
    """Returns the token for the user that logs in"""
    return auth_controllers.login_controller(
        email=form_data.username,
        password=form_data.password,
        session=session,
        response=response,
    )


@router.post("/signup", response_model=UserRead, status_code=status.HTTP_201_CREATED)
def signup(
    signup_data: InvitationSignupRequest,
    session: Session = Depends(get_session),
):
    """Returns the new user that signs up."""
    return auth_controllers.signup_controller(
        session=session,
        data=signup_data,
    )


@router.post("/refresh", response_model=Token, status_code=status.HTTP_200_OK)
def refresh_token(
    request: Request,
    response: Response,
    body: Optional[RefreshTokenRequest] = None,
    session: Session = Depends(get_session),
):
    """Rotates the refresh token and returns a new access/refresh token pair."""
    return auth_controllers.refresh_token_controller(
        session=session,
        request=request,
        response=response,
        body=body,
    )


@router.post("/logout", status_code=status.HTTP_200_OK)
def logout_user(
    request: Request,
    response: Response,
    body: Optional[RefreshTokenRequest] = None,
    session: Session = Depends(get_session),
) -> dict[str, str]:
    """Logs a user out."""
    return auth_controllers.logout_controller(
        session=session,
        request=request,
        response=response,
        body=body,
    )


@router.get("/invitations", response_model=list[InvitationRead], status_code=status.HTTP_200_OK)
def list_invitations(
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_admin),
):
    """Lists invitations."""
    return auth_controllers.list_invitations_controller(
        session=session,
        current_user=current_user,
    )


@router.post("/invitations/staff", response_model=InvitationCreatedResponse, status_code=status.HTTP_201_CREATED)
def create_staff_invitation(
    invitation_data: StaffInvitationCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Creates a staff invitation."""
    return auth_controllers.create_staff_invitation_controller(
        current_user=current_user,
        session=session,
        data=invitation_data,
    )


@router.post("/invitations/provision", response_model=DepartmentProvisioningResponse, status_code=status.HTTP_201_CREATED)
def provision_department_and_admin_invitation(
    provision_data: DepartmentAdminProvisioningCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_super_admin),
):
    """Super admin provisions department and simultaneously dept admin invitation."""
    return auth_controllers.provision_department_controller(
        session=session,
        current_user=current_user,
        data=provision_data,
    )


@router.post("/invitations/admin", response_model=InvitationCreatedResponse, status_code=status.HTTP_201_CREATED)
def provision_department_admin_invitation(
    invitation_data: DepartmentAdminInvitationCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_super_admin),
):
    """Super admin creates an invitation for a department admin for an existing department."""
    return auth_controllers.create_department_admin_invitation_controller(
        session=session,
        current_user=current_user,
        data=invitation_data,
    )


@router.post("/invitations/{invitation_id}/revoke", response_model=InvitationRead, status_code=status.HTTP_200_OK)
def revoke_invitation(
    invitation_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_admin),
):
    """Revokes an invitation."""
    return auth_controllers.revoke_invitation_controller(
        session=session,
        current_user=current_user,
        invitation_id=invitation_id,
    )
