from msal import ConfidentialClientApplication
import requests
import os

from dotenv import load_dotenv
from datetime import datetime

load_dotenv()


# ==========================================
# Azure Configuration
# ==========================================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OBJECT_ID = os.getenv("OBJECT_ID")


# ==========================================
# Generate Microsoft Graph Access Token
# ==========================================

def get_access_token():

    authority = (
        f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    token_result = app.acquire_token_for_client(
        scopes=[
            "https://graph.microsoft.com/.default"
        ]
    )

    if "access_token" not in token_result:

        return None

    return token_result["access_token"]


# ==========================================
# Create Teams Meeting
# ==========================================
def get_online_meeting_details(access_token, online_meeting_id, join_web_url=None):
    """Get Teams join URL, meeting ID and passcode."""
    url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/onlineMeetings/{online_meeting_id}"
    )
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"$select": "id,subject,joinWebUrl,joinMeetingIdSettings"}
    response = requests.get(url, headers=headers, params=params)
    try:
        data = response.json()
    except Exception:
        data = {"message": response.text}
    if response.status_code != 200:
        return {"status": "Failed", "statusCode": response.status_code, "response": data}
    settings = data.get("joinMeetingIdSettings", {})
    return {
        "status": "Success",
        "onlineMeetingId": data.get("id"),
        "joinWebUrl": data.get("joinWebUrl") or join_web_url,
        "meetingId": settings.get("joinMeetingId"),
        "passcode": settings.get("passcode")
    }


def update_event_invitation_body(access_token, event_id, candidate_name,
                                 subject, start_date_time, end_date_time,
                                 join_web_url, meeting_id, passcode,
                                 interviewers):
    """Add the custom HR invitation while preserving the Teams body blob."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    event_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/events/{event_id}"
    )

    # Microsoft recommends fetching and preserving the existing body for
    # an online meeting so the Teams meeting blob is not removed.
    get_response = requests.get(
        event_url,
        headers=headers,
        params={"$select": "body"}
    )
    try:
        existing_event = get_response.json()
    except Exception:
        existing_event = {"message": get_response.text}
    if get_response.status_code != 200:
        return {
            "status": "Failed",
            "statusCode": get_response.status_code,
            "response": existing_event
        }

    existing_body = existing_event.get("body", {})
    existing_content = existing_body.get("content", "")

    interview_date = start_date_time[:10]
    start_time = start_date_time[11:16]
    end_time = end_date_time[11:16]

    interviewer_text = ""
    if interviewers:
        interviewer_text = (
            "<p><strong>Interviewer(s):</strong> "
            + ", ".join(interviewers)
            + "</p>"
        )

    meeting_details = ""
    if meeting_id:
        meeting_details += f"<p><strong>Meeting ID:</strong> {meeting_id}</p>"
    if passcode:
        meeting_details += f"<p><strong>Passcode:</strong> {passcode}</p>"

    # One calendar event has one shared body for all attendees. Therefore the
    # wording is intentionally neutral rather than addressing only one person.
    custom_body = (
        "<html><body>"
        "<p>Hello,</p>"
        "<p>You have been invited to attend/conduct an interview for the following candidate.</p>"
        f"<p><strong>Candidate Name:</strong> {candidate_name}<br>"
        f"<strong>Interview Subject:</strong> {subject}<br>"
        f"<strong>Interview Date:</strong> {interview_date}<br>"
        f"<strong>Interview Time:</strong> {start_time} - {end_time} IST</p>"
        f"{interviewer_text}"
        "<hr>"
        "<h3>Microsoft Teams Meeting</h3>"
        f"<p><a href=\"{join_web_url}\">"
        "Click here to join the Microsoft Teams Interview"
        "</a></p>"
        f"{meeting_details}"
        "<p>Please join the meeting at the scheduled time and conduct the interview with the candidate.</p>"
        "<p>Regards,<br>HR Recruitment Team</p>"
        # Preserve the Teams-generated meeting content/blob.
        f"{existing_content}"
        "</body></html>"
    )

    patch_response = requests.patch(
        event_url,
        headers=headers,
        json={"body": {"contentType": "HTML", "content": custom_body}}
    )
    try:
        patch_data = patch_response.json()
    except Exception:
        patch_data = {"message": patch_response.text}

    if patch_response.status_code not in (200, 201):
        return {
            "status": "Failed",
            "statusCode": patch_response.status_code,
            "response": patch_data
        }
    return {"status": "Success"}


def create_teams_meeting(candidateName, email, startDateTime, endDateTime,
                         subject, interviewers, interviewersEmail):
    access_token = get_access_token()
    if access_token is None:
        return {"status": "Failed", "message": "Unable to generate access token."}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    teams_meeting_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/events"
    )

    attendees = [{
        "emailAddress": {"address": email, "name": candidateName},
        "type": "required"
    }]
    for interviewer_name, interviewer_email in zip(interviewers, interviewersEmail):
        attendees.append({
            "emailAddress": {"address": interviewer_email, "name": interviewer_name},
            "type": "required"
        })

    # First create the calendar-backed Teams event. Teams details are generated
    # by Graph only after this event is created.
    teams_meeting_payload = {
        "subject": f"{subject} - {candidateName}",
        "body": {
            "contentType": "HTML",
            "content": (
                f"<html><body><p>Interview for <strong>{candidateName}</strong>.</p>"
                f"<p><strong>Interview Type:</strong> {subject}</p></body></html>"
            )
        },
        "start": {"dateTime": startDateTime, "timeZone": "India Standard Time"},
        "end": {"dateTime": endDateTime, "timeZone": "India Standard Time"},
        "location": {"displayName": "Microsoft Teams"},
        "attendees": attendees,
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
        "responseRequested": True,
        "allowNewTimeProposals": True,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": 30,
        "showAs": "busy",
        "importance": "normal",
        "sensitivity": "normal"
    }

    response = requests.post(
        teams_meeting_url,
        headers=headers,
        json=teams_meeting_payload
    )
    try:
        response_data = response.json()
    except Exception:
        return {"status": "Failed", "message": response.text}

    if response.status_code != 201:
        return {"status": "Failed", "statusCode": response.status_code, "response": response_data}

    event_id = response_data.get("id")
    online_meeting = response_data.get("onlineMeeting", {})
    join_web_url = online_meeting.get("joinUrl")
    online_meeting_id = online_meeting.get("id")

    # Some event responses expose only joinUrl. Resolve the meeting ID if needed.
    if not online_meeting_id and join_web_url:
        lookup_url = (
            "https://graph.microsoft.com/v1.0/"
            f"users/{OBJECT_ID}/onlineMeetings"
        )
        lookup_response = requests.get(
            lookup_url,
            headers=headers,
            params={"$filter": f"JoinWebUrl eq '{join_web_url}'"}
        )
        try:
            lookup_data = lookup_response.json()
        except Exception:
            lookup_data = {}
        if lookup_response.status_code == 200 and lookup_data.get("value"):
            online_meeting_id = lookup_data["value"][0].get("id")

    meeting_id = None
    passcode = None
    if online_meeting_id:
        details = get_online_meeting_details(
            access_token, online_meeting_id, join_web_url
        )
        if details.get("status") == "Success":
            join_web_url = details.get("joinWebUrl") or join_web_url
            meeting_id = details.get("meetingId")
            passcode = details.get("passcode")

    # Add the custom HR email body after Graph has generated the Teams details.
    body_update = update_event_invitation_body(
        access_token=access_token,
        event_id=event_id,
        candidate_name=candidateName,
        subject=subject,
        start_date_time=startDateTime,
        end_date_time=endDateTime,
        join_web_url=join_web_url,
        meeting_id=meeting_id,
        passcode=passcode,
        interviewers=interviewers
    )

    if body_update.get("status") != "Success":
        return {
            "status": "Failed",
            "message": "Meeting was created, but the custom invitation body could not be updated.",
            "eventId": event_id,
            "onlineMeetingId": online_meeting_id,
            "joinWebUrl": join_web_url,
            "bodyUpdate": body_update
        }

    start = datetime.fromisoformat(startDateTime.replace("Z", "+00:00"))
    end = datetime.fromisoformat(endDateTime.replace("Z", "+00:00"))

    return {
        "status": "Success",
        "candidateName": candidateName,
        "candidateEmail": email,
        "interviewers": interviewers,
        "interviewersEmail": interviewersEmail,
        "eventId": event_id,
        "joinWebUrl": join_web_url,
        "onlineMeetingId": online_meeting_id,
        "meetingId": meeting_id,
        "passcode": passcode,
        "startDateTime": startDateTime,
        "endDateTime": endDateTime,
        "subject": response_data.get("subject"),
        "duration": str(end - start)
    }


# ==========================================
# Get Online Meeting from Calendar Event
# ==========================================
def get_online_meeting_from_event(
    access_token,
    event_id
):
    """
    Resolve a calendar event to its Teams onlineMeetingId.

    Flow:
        Event ID -> calendar event -> joinUrl
        -> JoinWebUrl lookup -> onlineMeetingId
    """

    event_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/events/{event_id}"
    )

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    params = {
        "$select": "id,subject,onlineMeeting"
    }

    response = requests.get(
        event_url,
        headers=headers,
        params=params
    )

    try:
        response_data = response.json()
    except Exception:
        response_data = {"message": response.text}

    if response.status_code != 200:
        return {
            "status": "Failed",
            "statusCode": response.status_code,
            "response": response_data
        }

    online_meeting = response_data.get("onlineMeeting", {})
    join_url = online_meeting.get("joinUrl")

    if not join_url:
        return {
            "status": "Failed",
            "message": (
                "The calendar event does not contain "
                "a Teams online meeting join URL."
            ),
            "eventId": event_id
        }

    meetings_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/onlineMeetings"
    )

    meeting_params = {
        "$filter": f"JoinWebUrl eq '{join_url}'"
    }

    meeting_response = requests.get(
        meetings_url,
        headers=headers,
        params=meeting_params
    )

    try:
        meeting_data = meeting_response.json()
    except Exception:
        meeting_data = {"message": meeting_response.text}

    if meeting_response.status_code != 200:
        return {
            "status": "Failed",
            "statusCode": meeting_response.status_code,
            "response": meeting_data
        }

    meetings = meeting_data.get("value", [])

    if not meetings:
        return {
            "status": "Failed",
            "message": (
                "No Teams online meeting was found for "
                "the calendar event join URL."
            ),
            "eventId": event_id,
            "joinWebUrl": join_url
        }

    meeting = meetings[0]

    return {
        "status": "Success",
        "eventId": event_id,
        "onlineMeetingId": meeting.get("id"),
        "joinWebUrl": join_url,
        "subject": meeting.get("subject")
    }


# ==========================================
# Get Transcripts for a Scheduled Meeting
# ==========================================
def get_transcripts(
    access_token,
    online_meeting_id
):
    """Return transcript metadata for a Teams meeting."""

    transcripts_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/onlineMeetings/"
        f"{online_meeting_id}/transcripts"
    )

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(
        transcripts_url,
        headers=headers
    )

    try:
        response_data = response.json()
    except Exception:
        response_data = {"message": response.text}

    if response.status_code != 200:
        return {
            "status": "Failed",
            "statusCode": response.status_code,
            "response": response_data
        }

    return {
        "status": "Success",
        "transcripts": response_data.get("value", [])
    }


# ==========================================
# Get Actual Transcript Content
# ==========================================
def get_transcript_content(
    access_token,
    online_meeting_id,
    transcript_id
):
    """Retrieve actual transcript content in VTT format."""

    transcript_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/onlineMeetings/"
        f"{online_meeting_id}/transcripts/"
        f"{transcript_id}/content"
    )

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "text/vtt"
    }

    response = requests.get(
        transcript_url,
        headers=headers
    )

    if response.status_code != 200:
        try:
            response_data = response.json()
        except Exception:
            response_data = {"message": response.text}

        return {
            "status": "Failed",
            "statusCode": response.status_code,
            "response": response_data
        }

    return {
        "status": "Success",
        "transcriptId": transcript_id,
        "onlineMeetingId": online_meeting_id,
        "content": response.text
    }


# ==========================================
# Get Latest Available Transcript by Event ID
# ==========================================
def get_latest_transcript(
    event_id
):
    """
    End-to-end flow:
        Event ID
        -> joinUrl
        -> onlineMeetingId
        -> list transcripts
        -> latest transcript
        -> transcript content

    If no transcript is available, return NotAvailable.
    """

    access_token = get_access_token()

    if access_token is None:
        return {
            "status": "Failed",
            "message": "Unable to generate access token."
        }

    # Step 1: Event ID -> onlineMeetingId
    meeting_result = get_online_meeting_from_event(
        access_token,
        event_id
    )

    if meeting_result.get("status") != "Success":
        return meeting_result

    online_meeting_id = meeting_result.get("onlineMeetingId")

    if not online_meeting_id:
        return {
            "status": "Failed",
            "message": (
                "No onlineMeetingId was returned for "
                "the calendar event."
            ),
            "eventId": event_id
        }

    # Step 2: Get transcript metadata
    transcript_result = get_transcripts(
        access_token,
        online_meeting_id
    )

    if transcript_result.get("status") != "Success":
        return transcript_result

    transcripts = transcript_result.get("transcripts", [])

    if not transcripts:
        return {
            "status": "NotAvailable",
            "message": "No transcription available.",
            "eventId": event_id,
            "onlineMeetingId": online_meeting_id
        }

    # Step 3: Select newest transcript
    transcripts.sort(
        key=lambda item: item.get("createdDateTime", ""),
        reverse=True
    )

    latest = transcripts[0]
    transcript_id = latest.get("id")

    if not transcript_id:
        return {
            "status": "Failed",
            "message": (
                "A transcript was found, but its ID "
                "was not returned by Microsoft Graph."
            ),
            "eventId": event_id,
            "onlineMeetingId": online_meeting_id
        }

    # Step 4: Get transcript content
    content_result = get_transcript_content(
        access_token,
        online_meeting_id,
        transcript_id
    )

    if content_result.get("status") != "Success":
        return content_result

    return {
        "status": "Success",
        "eventId": event_id,
        "onlineMeetingId": online_meeting_id,
        "transcriptId": transcript_id,
        "createdDateTime": latest.get("createdDateTime"),
        "transcript": content_result.get("content")
    }

