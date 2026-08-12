from fastapi import FastAPI
from pydantic import BaseModel
from typing import List

from main_seperate_mail import create_teams_meeting, get_latest_transcript


app = FastAPI()



class TeamsMeetingRequest(BaseModel):

    candidateName: str
    email: str
    startDateTime: str
    endDateTime: str
    subject: str

    interviewers: List[str]
    interviewersEmail: List[str]



@app.get("/")
def home():

    return {
        "message": "Teams Meeting API is Running"
    }



@app.post("/create-teams-meeting")
def execute(request: TeamsMeetingRequest):

    return create_teams_meeting(

        request.candidateName,

        request.email,

        request.startDateTime,

        request.endDateTime,

        request.subject,

        request.interviewers,

        request.interviewersEmail
    )

# ==========================================
# Get Latest Teams Meeting Transcript
# ==========================================
class TeamsTranscriptRequest(BaseModel):

    eventId: str


@app.post("/get-transcript")
def get_transcript(request: TeamsTranscriptRequest):

    return get_latest_transcript(
        request.eventId
    )
