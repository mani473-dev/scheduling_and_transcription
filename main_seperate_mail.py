
from msal import ConfidentialClientApplication
import requests
import os

from dotenv import load_dotenv
from datetime import datetime


load_dotenv()


# ============================================================
# Azure Configuration
# ============================================================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OBJECT_ID = os.getenv("OBJECT_ID")


# ============================================================
# Generate Microsoft Graph Access Token
# ============================================================

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

        print("Unable to generate access token.")
        print(token_result)

        return None

    return token_result["access_token"]


# ============================================================
# Send Email
# ============================================================

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

    print(f"sendmail - {response.status_code}")
    print("-----------------------------------------------")
    print(url)

    if response.status_code not in [200, 202]:

        print("Email sending failed.")
        print(response.text)

    return response.status_code


# ============================================================
# Create Teams Meeting
# ============================================================

def create_teams_meeting(
    candidateName,
    email,
    startDateTime,
    endDateTime,
    subject,
    interviewers,
    interviewersEmail
):

    # ========================================================
    # Generate Access Token
    # ========================================================

    access_token = get_access_token()

    if access_token is None:

        return {
            "status": "Failed",
            "message": "Unable to generate access token."
        }


    # ========================================================
    # Normalize DateTime
    # ========================================================

    # Expected input:
    #
    # 2026-09-08T14:30:00
    #
    # or:
    #
    # 2026-09-08T14:30:00Z
    #
    # We are treating the supplied time as IST.
    #
    # Therefore we remove Z before sending it to Graph
    # and explicitly specify:
    #
    # India Standard Time
    # ========================================================

    graph_start_datetime = startDateTime.replace(
        "Z",
        ""
    )

    graph_end_datetime = endDateTime.replace(
        "Z",
        ""
    )


    # ========================================================
    # Format Interview Date & Time for Email
    # ========================================================

    interview_date = graph_start_datetime[:10]

    interview_start_time = graph_start_datetime[11:16]

    interview_end_time = graph_end_datetime[11:16]

    interview_time = (
        f"{interview_start_time} - "
        f"{interview_end_time} IST"
    )


    # ========================================================
    # Request Headers
    # ========================================================

    headers = {

        "Authorization": (
            f"Bearer {access_token}"
        ),

        "Content-Type": "application/json"
    }


    # ========================================================
    # Microsoft Graph Calendar Event URL
    # ========================================================

    teams_meeting_url = (
        "https://graph.microsoft.com/v1.0/"
        f"users/{OBJECT_ID}/events"
    )


    # ========================================================
    # Build Interviewer Attendees
    # ========================================================

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


    # ========================================================
    # Calendar-backed Teams Meeting Payload
    # ========================================================
    #
    # IMPORTANT:
    #
    # start.dateTime = local Indian time
    #
    # start.timeZone = India Standard Time
    #
    # end.dateTime = local Indian time
    #
    # end.timeZone = India Standard Time
    #
    # This prevents Microsoft Graph from interpreting
    # 14:30 as UTC or another timezone.
    #
    # ========================================================

    teams_meeting_payload = {

        "subject": subject,

        "body": {

            "contentType": "HTML",

            "content": (
                "<html><body>"
                "<p>"
                "Interview meeting created by "
                "HR Recruitment Team."
                "</p>"
                "</body></html>"
            )
        },

        "start": {

            "dateTime": graph_start_datetime,

            "timeZone": "India Standard Time"
        },

        "end": {

            "dateTime": graph_end_datetime,

            "timeZone": "India Standard Time"
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


    # ========================================================
    # Print Payload for Debugging
    # ========================================================

    print("")
    print("===============================================")
    print("TEAMS MEETING PAYLOAD")
    print("===============================================")

    print(teams_meeting_payload)

    print("===============================================")
    print("")


    # ========================================================
    # Create Calendar Event / Teams Meeting
    # ========================================================

    response = requests.post(

        teams_meeting_url,

        headers=headers,

        json=teams_meeting_payload
    )


    # ========================================================
    # Read Microsoft Graph Response
    # ========================================================

    try:

        response_data = response.json()

    except Exception:

        return {

            "status": "Failed",

            "statusCode": response.status_code,

            "message": response.text
        }


    # ========================================================
    # Check Meeting Creation
    # ========================================================

    if response.status_code != 201:

        print("")
        print("===============================================")
        print("MICROSOFT GRAPH ERROR")
        print("===============================================")
        print(response_data)
        print("===============================================")
        print("")

        return {

            "status": "Failed",

            "statusCode": response.status_code,

            "response": response_data
        }


    # ========================================================
    # Meeting Created Successfully
    # ========================================================

    print("")
    print("===============================================")
    print("TEAMS MEETING CREATED")
    print("===============================================")


    # ========================================================
    # Calculate Duration
    # ========================================================

    start = datetime.fromisoformat(
        graph_start_datetime
    )

    end = datetime.fromisoformat(
        graph_end_datetime
    )

    duration = end - start


    # ========================================================
    # Get Calendar Event ID
    # ========================================================

    event_id = response_data.get(
        "id"
    )


    # ========================================================
    # Get Teams Online Meeting
    # ========================================================

    online_meeting = response_data.get(
        "onlineMeeting",
        {}
    )

    join_web_url = online_meeting.get(
        "joinUrl"
    )

    online_meeting_id = online_meeting.get(
        "id"
    )


    # ========================================================
    # Debug Meeting Information
    # ========================================================

    print("Event ID:")
    print(event_id)

    print("Teams Join URL:")
    print(join_web_url)

    print("Online Meeting ID:")
    print(online_meeting_id)

    print("Start:")
    print(graph_start_datetime)

    print("End:")
    print(graph_end_datetime)

    print("Duration:")
    print(duration)

    print("===============================================")
    print("")


    # ========================================================
    # Candidate Email
    # ========================================================

    candidate_body = f"""
<html>
<body>

<p>Dear {candidateName},</p>

<p>
We are pleased to inform you that your interview
has been scheduled.
</p>

<p>

<strong>Interview Subject:</strong>
{subject}
<br>

<strong>Interview Date:</strong>
{interview_date}
<br>

<strong>Interview Time:</strong>
{interview_time}

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

Regards,
<br>
HR Recruitment Team

</p>

</body>
</html>
"""


    # ========================================================
    # Send Candidate Email
    # ========================================================

    candidate_mail_status = send_mail(

        access_token,

        email,

        f"{subject} - {candidateName}",

        candidate_body
    )


    # ========================================================
    # Interviewer Emails
    # ========================================================

    interviewer_mail_status = []


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

<strong>Candidate Name:</strong>
{candidateName}
<br>

<strong>Interview Subject:</strong>
{subject}
<br>

<strong>Interview Date:</strong>
{interview_date}
<br>

<strong>Interview Time:</strong>
{interview_time}

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

Regards,
<br>
HR Recruitment Team

</p>

</body>
</html>
"""


        mail_status = send_mail(

            access_token,

            interviewer_email,

            f"Interview Assignment - {candidateName}",

            interviewer_body
        )


        interviewer_mail_status.append({

            "interviewer": interviewer_name,

            "email": interviewer_email,

            "statusCode": mail_status
        })


    # ========================================================
    # Final Response
    #
    # IMPORTANT:
    #
    # This return is OUTSIDE the interviewer loop.
    #
    # Therefore all interviewers get their emails.
    # ========================================================

    return {

        "status": "Success",

        "candidateName": candidateName,

        "candidateEmail": email,

        "interviewers": interviewers,

        "interviewersEmail": interviewersEmail,

        "eventId": event_id,

        "joinWebUrl": join_web_url,

        "onlineMeetingId": online_meeting_id,

        "startDateTime": graph_start_datetime,

        "endDateTime": graph_end_datetime,

        "timeZone": "India Standard Time",

        "subject": response_data.get(
            "subject"
        ),

        "duration": str(duration),

        "candidateMailStatus": candidate_mail_status,

        "interviewerMailStatus": interviewer_mail_status
    }


# ============================================================
# Get Online Meeting from Calendar Event
# ============================================================

def get_online_meeting_from_event(
    access_token,
    event_id
):

    """
    Flow:

        Event ID
            ↓
        Calendar Event
            ↓
        onlineMeeting.joinUrl
            ↓
        Search onlineMeetings
            ↓
        onlineMeetingId
    """


    # ========================================================
    # Get Calendar Event
    # ========================================================

    event_url = (

        "https://graph.microsoft.com/v1.0/"

        f"users/{OBJECT_ID}/events/{event_id}"

    )


    headers = {

        "Authorization": (
            f"Bearer {access_token}"
        )
    }


    params = {

        "$select": (
            "id,"
            "subject,"
            "onlineMeeting,"
            "start,"
            "end"
        )
    }


    response = requests.get(

        event_url,

        headers=headers,

        params=params
    )


    # ========================================================
    # Parse Response
    # ========================================================

    try:

        response_data = response.json()

    except Exception:

        response_data = {

            "message": response.text
        }


    # ========================================================
    # Check Response
    # ========================================================

    if response.status_code != 200:

        return {

            "status": "Failed",

            "statusCode": response.status_code,

            "response": response_data
        }


    # ========================================================
    # Get Online Meeting
    # ========================================================

    online_meeting = response_data.get(

        "onlineMeeting",

        {}
    )


    join_url = online_meeting.get(

        "joinUrl"
    )


    # ========================================================
    # No Join URL
    # ========================================================

    if not join_url:

        return {

            "status": "Failed",

            "message": (
                "The calendar event does not contain "
                "a Teams online meeting join URL."
            ),

            "eventId": event_id
        }


    # ========================================================
    # Get Online Meeting ID
    # ========================================================

    meetings_url = (

        "https://graph.microsoft.com/v1.0/"

        f"users/{OBJECT_ID}/onlineMeetings"

    )


    meeting_params = {

        "$filter": (
            f"JoinWebUrl eq '{join_url}'"
        )
    }


    meeting_response = requests.get(

        meetings_url,

        headers=headers,

        params=meeting_params
    )


    # ========================================================
    # Parse Meeting Response
    # ========================================================

    try:

        meeting_data = meeting_response.json()

    except Exception:

        meeting_data = {

            "message": meeting_response.text
        }


    # ========================================================
    # Check Meeting Response
    # ========================================================

    if meeting_response.status_code != 200:

        return {

            "status": "Failed",

            "statusCode": meeting_response.status_code,

            "response": meeting_data
        }


    meetings = meeting_data.get(

        "value",

        []
    )


    # ========================================================
    # Meeting Not Found
    # ========================================================

    if not meetings:

        return {

            "status": "Failed",

            "message": (
                "No Teams online meeting was found "
                "for the calendar event join URL."
            ),

            "eventId": event_id,

            "joinWebUrl": join_url
        }


    # ========================================================
    # Get First Meeting
    # ========================================================

    meeting = meetings[0]


    # ========================================================
    # Return Meeting Information
    # ========================================================

    return {

        "status": "Success",

        "eventId": event_id,

        "onlineMeetingId": meeting.get(
            "id"
        ),

        "joinWebUrl": join_url,

        "subject": meeting.get(
            "subject"
        )
    }


# ============================================================
# Get Transcripts
# ============================================================

def get_transcripts(

    access_token,

    online_meeting_id

):

    """
    Return transcript metadata
    for a Teams meeting.
    """


    transcripts_url = (

        "https://graph.microsoft.com/v1.0/"

        f"users/{OBJECT_ID}/onlineMeetings/"

        f"{online_meeting_id}/transcripts"

    )


    headers = {

        "Authorization": (
            f"Bearer {access_token}"
        )
    }


    response = requests.get(

        transcripts_url,

        headers=headers
    )


    # ========================================================
    # Parse Response
    # ========================================================

    try:

        response_data = response.json()

    except Exception:

        response_data = {

            "message": response.text
        }


    # ========================================================
    # Check Response
    # ========================================================

    if response.status_code != 200:

        return {

            "status": "Failed",

            "statusCode": response.status_code,

            "response": response_data
        }


    # ========================================================
    # Return Transcript Metadata
    # ========================================================

    return {

        "status": "Success",

        "transcripts": response_data.get(
            "value",
            []
        )
    }


# ============================================================
# Get Actual Transcript Content
# ============================================================

def get_transcript_content(

    access_token,

    online_meeting_id,

    transcript_id

):

    """
    Retrieve actual transcript content
    in VTT format.
    """


    transcript_url = (

        "https://graph.microsoft.com/v1.0/"

        f"users/{OBJECT_ID}/onlineMeetings/"

        f"{online_meeting_id}/transcripts/"

        f"{transcript_id}/content"

    )


    headers = {

        "Authorization": (
            f"Bearer {access_token}"
        ),

        "Accept": "text/vtt"
    }


    response = requests.get(

        transcript_url,

        headers=headers
    )


    # ========================================================
    # Check Response
    # ========================================================

    if response.status_code != 200:

        try:

            response_data = response.json()

        except Exception:

            response_data = {

                "message": response.text
            }


        return {

            "status": "Failed",

            "statusCode": response.status_code,

            "response": response_data
        }


    # ========================================================
    # Return Transcript
    # ========================================================

    return {

        "status": "Success",

        "transcriptId": transcript_id,

        "onlineMeetingId": online_meeting_id,

        "content": response.text
    }


# ============================================================
# Get Latest Available Transcript by Event ID
# ============================================================

def get_latest_transcript(

    event_id

):

    """
    End-to-end flow:

        Event ID
            ↓
        Calendar Event
            ↓
        Join URL
            ↓
        Online Meeting ID
            ↓
        List Transcripts
            ↓
        Latest Transcript
            ↓
        Transcript Content

    If no transcript is available,
    return NotAvailable.
    """


    # ========================================================
    # Generate Access Token
    # ========================================================

    access_token = get_access_token()


    if access_token is None:

        return {

            "status": "Failed",

            "message": (
                "Unable to generate access token."
            )
        }


    # ========================================================
    # Step 1
    #
    # Event ID -> Online Meeting ID
    # ========================================================

    meeting_result = get_online_meeting_from_event(

        access_token,

        event_id
    )


    if meeting_result.get(
        "status"
    ) != "Success":

        return meeting_result


    online_meeting_id = meeting_result.get(

        "onlineMeetingId"
    )


    # ========================================================
    # Check Online Meeting ID
    # ========================================================

    if not online_meeting_id:

        return {

            "status": "Failed",

            "message": (
                "No Teams online meeting was found "
                "for the calendar event."
            ),

            "eventId": event_id
        }


    # ========================================================
    # Step 2
    #
    # Get Transcript Metadata
    # ========================================================

    transcript_result = get_transcripts(

        access_token,

        online_meeting_id
    )


    if transcript_result.get(
        "status"
    ) != "Success":

        return transcript_result


    transcripts = transcript_result.get(

        "transcripts",

        []
    )


    # ========================================================
    # No Transcript
    # ========================================================

    if not transcripts:

        return {

            "status": "NotAvailable",

            "message": (
                "No transcription available."
            ),

            "eventId": event_id,

            "onlineMeetingId": online_meeting_id
        }


    # ========================================================
    # Step 3
    #
    # Select Newest Transcript
    # ========================================================

    transcripts.sort(

        key=lambda item:
        item.get(
            "createdDateTime",
            ""
        ),

        reverse=True
    )


    latest = transcripts[0]


    transcript_id = latest.get(

        "id"
    )


    # ========================================================
    # Check Transcript ID
    # ========================================================

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


    # ========================================================
    # Step 4
    #
    # Get Actual Transcript Content
    # ========================================================

    content_result = get_transcript_content(

        access_token,

        online_meeting_id,

        transcript_id
    )


    if content_result.get(
        "status"
    ) != "Success":

        return content_result


    # ========================================================
    # Final Transcript Response
    # ========================================================

    return {

        "status": "Success",

        "eventId": event_id,

        "onlineMeetingId": online_meeting_id,

        "transcriptId": transcript_id,

        "createdDateTime": latest.get(
            "createdDateTime"
        ),

        "transcript": content_result.get(
            "content"
        )
    }
