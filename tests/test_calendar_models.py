"""Tests for outlookctl calendar models."""

import pytest
from outlookctl.models import (
    EventId,
    Attendee,
    RecurrenceInfo,
    EventSummary,
    EventDetail,
    CalendarListResult,
    EventCreateResult,
    EventSendResult,
    EventRespondResult,
)


class TestEventId:
    def test_to_dict(self):
        event_id = EventId(entry_id="event123", store_id="store456")
        result = event_id.to_dict()
        assert result == {"entry_id": "event123", "store_id": "store456"}


class TestAttendee:
    def test_to_dict(self):
        attendee = Attendee(
            name="John Doe",
            email="john@example.com",
            type="required",
            response="accepted",
        )
        result = attendee.to_dict()
        assert result == {
            "name": "John Doe",
            "email": "john@example.com",
            "type": "required",
            "response": "accepted",
        }


class TestRecurrenceInfo:
    def test_to_dict_minimal(self):
        recurrence = RecurrenceInfo(type="daily", interval=1)
        result = recurrence.to_dict()
        assert result == {"type": "daily", "interval": 1}

    def test_to_dict_weekly(self):
        recurrence = RecurrenceInfo(
            type="weekly",
            interval=1,
            days_of_week=["monday", "wednesday"],
            end_date="2025-12-31",
        )
        result = recurrence.to_dict()
        assert result["type"] == "weekly"
        assert result["days_of_week"] == ["monday", "wednesday"]
        assert result["end_date"] == "2025-12-31"

    def test_to_dict_monthly(self):
        recurrence = RecurrenceInfo(
            type="monthly",
            interval=1,
            day_of_month=15,
            occurrences=12,
        )
        result = recurrence.to_dict()
        assert result["day_of_month"] == 15
        assert result["occurrences"] == 12


class TestEventSummary:
    def test_to_dict(self):
        summary = EventSummary(
            id=EventId(entry_id="e1", store_id="s1"),
            subject="Team Meeting",
            start="2025-01-20T10:00:00",
            end="2025-01-20T11:00:00",
            location="Conference Room A",
            organizer="organizer@example.com",
            is_recurring=False,
            is_all_day=False,
            is_meeting=True,
            response_status="accepted",
            busy_status="busy",
        )
        result = summary.to_dict()
        assert result["subject"] == "Team Meeting"
        assert result["location"] == "Conference Room A"
        assert result["is_meeting"] is True
        assert result["response_status"] == "accepted"


class TestEventDetail:
    def test_to_dict_minimal(self):
        detail = EventDetail(
            id=EventId(entry_id="e1", store_id="s1"),
            subject="Meeting",
            start="2025-01-20T10:00:00",
            end="2025-01-20T11:00:00",
            location="",
            organizer="org@example.com",
            is_recurring=False,
            is_all_day=False,
            is_meeting=False,
            response_status="organizer",
            busy_status="busy",
        )
        result = detail.to_dict()
        assert result["subject"] == "Meeting"
        assert "body" not in result
        assert "attendees" not in result
        assert "recurrence" not in result

    def test_to_dict_with_attendees(self):
        detail = EventDetail(
            id=EventId(entry_id="e1", store_id="s1"),
            subject="Team Sync",
            start="2025-01-20T10:00:00",
            end="2025-01-20T11:00:00",
            location="Room B",
            organizer="boss@example.com",
            is_recurring=False,
            is_all_day=False,
            is_meeting=True,
            response_status="organizer",
            busy_status="busy",
            body="Agenda: Discuss project status",
            attendees=[
                Attendee(
                    name="Alice",
                    email="alice@example.com",
                    type="required",
                    response="accepted",
                ),
                Attendee(
                    name="Bob",
                    email="bob@example.com",
                    type="optional",
                    response="tentative",
                ),
            ],
            categories=["Work", "Important"],
            reminder_minutes=15,
        )
        result = detail.to_dict()
        assert result["body"] == "Agenda: Discuss project status"
        assert len(result["attendees"]) == 2
        assert result["attendees"][0]["email"] == "alice@example.com"
        assert result["categories"] == ["Work", "Important"]
        assert result["reminder_minutes"] == 15

    def test_to_dict_with_recurrence(self):
        detail = EventDetail(
            id=EventId(entry_id="e1", store_id="s1"),
            subject="Weekly Standup",
            start="2025-01-20T09:00:00",
            end="2025-01-20T09:30:00",
            location="",
            organizer="team@example.com",
            is_recurring=True,
            is_all_day=False,
            is_meeting=True,
            response_status="organizer",
            busy_status="busy",
            recurrence=RecurrenceInfo(
                type="weekly",
                interval=1,
                days_of_week=["monday"],
                end_date="2025-12-31",
            ),
        )
        result = detail.to_dict()
        assert result["is_recurring"] is True
        assert result["recurrence"]["type"] == "weekly"
        assert result["recurrence"]["days_of_week"] == ["monday"]


class TestCalendarListResult:
    def test_to_dict(self):
        result = CalendarListResult(
            calendar="Calendar",
            start_date="2025-01-20T00:00:00",
            end_date="2025-01-27T00:00:00",
            items=[],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["calendar"] == "Calendar"
        assert output["start_date"] == "2025-01-20T00:00:00"
        assert output["items"] == []


class TestEventCreateResult:
    def test_to_dict_draft(self):
        result = EventCreateResult(
            success=True,
            id=EventId(entry_id="event1", store_id="store1"),
            saved_to="Calendar",
            subject="New Meeting",
            start="2025-01-20T10:00:00",
            attendees=["alice@example.com", "bob@example.com"],
            is_draft=True,
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["id"]["entry_id"] == "event1"
        assert output["subject"] == "New Meeting"
        assert output["is_draft"] is True
        assert len(output["attendees"]) == 2

    def test_to_dict_sent(self):
        result = EventCreateResult(
            success=True,
            id=EventId(entry_id="event1", store_id="store1"),
            subject="Sent Meeting",
            is_draft=False,
        )
        output = result.to_dict()
        assert output["is_draft"] is False


class TestEventSendResult:
    def test_to_dict(self):
        result = EventSendResult(
            success=True,
            message="Meeting invitations sent",
            sent_at="2025-01-20T09:00:00",
            attendees=["alice@example.com", "bob@example.com"],
            subject="Team Meeting",
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["message"] == "Meeting invitations sent"
        assert output["sent_at"] == "2025-01-20T09:00:00"


class TestEventRespondResult:
    def test_to_dict_accept(self):
        result = EventRespondResult(
            success=True,
            response="accepted",
            subject="Team Meeting",
            organizer="boss@example.com",
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["response"] == "accepted"
        assert output["subject"] == "Team Meeting"

    def test_to_dict_decline(self):
        result = EventRespondResult(
            success=True,
            response="declined",
            subject="Optional Meeting",
        )
        output = result.to_dict()
        assert output["response"] == "declined"
