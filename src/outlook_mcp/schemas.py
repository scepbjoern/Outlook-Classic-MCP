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
EventAllDay = Annotated[bool, Field(
    strict=True,
    description="Native Outlook AllDayEvent. Use local dates or local midnight boundaries for all-day "
                "appointments. End is exclusive. Omit for a timed create or to preserve the type on update.",
)]
EventBusyStatus = Annotated[int, Field(
    strict=True, ge=0, le=4,
    description="Native Outlook availability (0=Free, 1=Tentative, 2=Busy, 3=Out of office, "
                "4=Working elsewhere). On create, omission selects Free with category P.AL or Busy otherwise. "
                "On update, omission preserves the stored value.",
)]
EventSensitivity = Annotated[int, Field(
    strict=True, ge=0, le=3,
    description="Native Outlook sensitivity (0=Normal, 1=Personal, 2=Private, 3=Confidential). "
                "Omit to keep the Outlook default on create or the current value on update.",
)]
EventScope = Annotated[Literal["occurrence", "series"] | None, Field(
    description="For recurring appointments select one occurrence or explicitly the entire series.",
)]
OriginalStart = Annotated[str | None, Field(
    description="Copy original_start (with UTC offset) from the listed occurrence, including moved exceptions.",
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
