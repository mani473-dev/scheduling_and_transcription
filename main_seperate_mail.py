#main_updatedBodyV1.py
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

# ==========================================
# Send Mail
# ==========================================

def send_mail(
    access_token,
    recipient_email,
    subject,
    body
):

    url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/sendMail"
    )

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    payload = {

        "message": {

            "subject": subject,

            "body": {
                "contentType": "HTML",
                "content": body
            },

            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient_email
                    }
                }
            ]
        },

        "saveToSentItems": True
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )
    print(f"sendmail- {response.status_code}")
    print("-----------------------------------------------")
    print(url)
    # print(payload)
    
    return response.status_code


def create_teams_meeting(
    candidateName,
    email,
    startDateTime,
    endDateTime,
    subject,
    interviewers,
    interviewersEmail
):



    # ==========================================
    # Generate Access Token
    # ==========================================

    access_token = get_access_token()

    if access_token is None:

        return {
            "status": "Failed",
            "message": "Unable to generate access token."
        }

    # ==========================================
    # Format Interview Date & Time
    # ==========================================

    interview_date = startDateTime[:10]

    interview_start_time = startDateTime[11:16]
    interview_end_time = endDateTime[11:16]

    interview_time = (
        f"{interview_start_time} - "
        f"{interview_end_time} IST"
    )


    # ==========================================
    # Request Headers
    # ==========================================

    headers = {

        "Authorization": (
            f"Bearer {access_token}"
        ),

        "Content-Type": "application/json"
    }


    # ==========================================
    # Microsoft Graph Calendar Event API
    # ==========================================
    # A calendar-backed Teams meeting is used so that
    # the meeting can later be associated with a
    # transcript through Microsoft Graph.

    teams_meeting_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/events"
    )



    # ==========================================
    # Build Interviewer Attendees
    # ==========================================

    attendees = []

    for interviewer_name, interviewer_email in zip(
        interviewers,
        interviewersEmail
    ):

        attendees.append({

            "emailAddress": {

                "address": interviewer_email,

                "name": interviewer_name
            },

            "type": "required"
        })


    # ==========================================
    # Calendar-backed Teams Meeting Payload
    # ==========================================
    # Creating the meeting as a calendar event with
    # isOnlineMeeting=True allows the meeting to be
    # associated with the organizer's calendar and
    # later used for transcript retrieval.
    #
    # startDateTime / endDateTime are expected in the
    # format:
    #   2026-08-10T14:30:00
    #
    # The timezone is explicitly specified below.
    # ==========================================

    teams_meeting_payload = calendar_events_url_payload = {

    "subject": f"{subject}",

    "body": {

        "contentType": "HTML",

        "content": (
            "<html><body>"
            "<p>Interview meeting created by HR Recruitment Team.</p>"
            "</body></html>"
        )
    },

    "location": {

        "displayName": "Microsoft Teams"
    },

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


    # ==========================================
    # Create Meeting
    # ==========================================

    response = requests.post(

        teams_meeting_url,

        headers=headers,

        json=teams_meeting_payload
    )


    # ==========================================
    # Read Microsoft Graph Response
    # ==========================================

    try:

        response_data = response.json()

    except Exception:

        return {

            "status": "Failed",

            "message": response.text
        }


    # ==========================================
    # Meeting Created Successfully
    # ==========================================

    if response.status_code == 201:

        # --------------------------------------
        # Calculate Duration
        # --------------------------------------

        start = datetime.fromisoformat(
            startDateTime.replace(
                "Z",
                "+00:00"
            )
        )

        end = datetime.fromisoformat(
            endDateTime.replace(
                "Z",
                "+00:00"
            )
        )

        duration = end - start


        # --------------------------------------
        # Get Calendar Event ID
        # --------------------------------------

        event_id = response_data.get(
            "id"
        )


        # --------------------------------------
        # Get Teams Meeting URL
        # --------------------------------------

        online_meeting = (
            response_data.get(
                "onlineMeeting",
                {}
            )
        )

        join_web_url = online_meeting.get(
            "joinUrl"
        )

        # ==========================================
    # Candidate Mail
    # ==========================================

    candidate_body = f"""
    <html>
    <body>

    <p>Dear {candidateName},</p>

    <p>
    We are pleased to inform you that your interview
    has been scheduled.
    </p>

    <p>
    <strong>Interview Subject:</strong> {subject}<br>
    <strong>Interview Date:</strong> {interview_date}<br>
    <strong>Interview Time:</strong> {interview_time}
    </p>

    <p>
    <a href="{join_web_url}">
    Join Microsoft Teams Meeting
    </a>
    </p>

    <p>
    Please join the meeting at the scheduled time.
    </p>

    <p>
    Regards,<br>
    HR Recruitment Team
    </p>

    </body>
    </html>
    """

    send_mail(
        access_token,
        email,
        f"{subject} - {candidateName}",
        candidate_body
    )

    # ==========================================
    # Interviewer Mails
    # ==========================================

    for interviewer_name, interviewer_email in zip(
        interviewers,
        interviewersEmail
    ):

        interviewer_body = f"""
        <html>
        <body>

        <p>Dear {interviewer_name},</p>

        <p>
        You have been assigned to conduct an interview
        for the following candidate.
        </p>

        <p>
        <strong>Candidate Name:</strong> {candidateName}<br>
        <strong>Interview Subject:</strong> {subject}<br>
        <strong>Interview Date:</strong> {interview_date}<br>
        <strong>Interview Time:</strong> {interview_time}
        </p>

        <p>
        <a href="{join_web_url}">
        Join Microsoft Teams Meeting
        </a>
        </p>

        <p>
        Please join the meeting at the scheduled time
        and conduct the interview with the candidate.
        </p>

        <p>
        Regards,<br>
        HR Recruitment Team
        </p>

        </body>
        </html>
        """

        send_mail(
            access_token,
            interviewer_email,
            f"Interview Assignment - {candidateName}",
            interviewer_body
        )


        # --------------------------------------
        # Return the calendar event ID.
        #
        # The event ID is important for tracking the
        # scheduled interview in your application.
        #
        # The onlineMeeting object contains the Teams
        # join URL.
        # --------------------------------------


        # --------------------------------------
        # Return Response
        # --------------------------------------

        return {

            "status": "Success",

            "candidateName": candidateName,

            "candidateEmail": email,

            "interviewers": interviewers,

            "interviewersEmail": interviewersEmail,

            "eventId": event_id,

            "joinWebUrl": join_web_url,

            "onlineMeetingId": online_meeting.get(
                "id"
            ),

            "startDateTime": startDateTime,

            "endDateTime": endDateTime,

            "subject": response_data.get(
                "subject"
            ),

            "duration": str(duration)
        }


    # ==========================================
    # Microsoft Graph Error
    # ==========================================

    return {

        "status": "Failed",

        "statusCode": response.status_code,

        "response": response_data
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
                "No Meeting started for the calendar event."
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
