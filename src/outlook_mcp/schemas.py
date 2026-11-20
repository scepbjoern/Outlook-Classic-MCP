"""Shared Pydantic models and enums for tool parameters."""

from __future__ import annotations

from enum import Enum
from typing import Annotated, Literal, Optional

from pydantic import BaseModel, Field


AccountSelector = Annotated[str | None, Field(
    description="Exact Outlook SMTP account. In a multi-account profile, provide account or store_id.",
)]
StoreSelector = Annotated[str | None, Field(
    description="Outlook StoreID from outlook_whoami or an item result. Must agree with account if both are supplied.",
)]


class ResponseFormat(str, Enum):
    JSON = "json"
    MARKDOWN = "markdown"


class Recurrence(BaseModel):
    """Recurrence pattern for calendar events.

    ``type`` is the cadence; ``interval`` is the spacing (every N days /
    weeks / months / years). Provide either ``occurrences`` or
    ``end_date`` to bound the series — if both are omitted the pattern
    runs indefinitely.
    """

    type: Literal["daily", "weekly", "monthly", "yearly"] = Field(
        description="Cadence: daily, weekly, monthly, or yearly."
    )
    interval: int = Field(
        default=1, ge=1, description="Every N units (days/weeks/months/years)."
    )
    occurrences: Optional[int] = Field(
        default=None, ge=1, description="End after N occurrences."
    )
    end_date: Optional[str] = Field(
        default=None,
        description="ISO-8601 date for series end (alternative to occurrences).",
    )
