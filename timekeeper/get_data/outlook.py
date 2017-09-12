import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # "6" refers to the index of a folder - in this case,
# the inbox.
inbox = inbox.Items
inbox = inbox.Items


class Message:
    def __init__(self, message):
        self.entry_id = message.EntryID
        self.sender_name = message.SenderName
        self.categories = message.Categories
        self.attachment_count = message.Attachments.Count
        self.sender_address = message.sender.address
        self.sent_to = message.To
        self.size = message.size
        self.cc = message.CC
        self.unread = message.UnRead
        self.recipient_names = [x.name for x in message.Recipients]
        self.recipient_addresses = [x.address for x in message.Recipients]
        self.received_time = message.ReceivedTime
        self.creation_time = message.CreationTime
        self.last_modification_time = message.LastModificationTime
        self.sent_on = message.SentOn
        self.subject = message.subject
        self.body = message.body
        self.body_most_recent_message = self.body.split('\r\nFrom:')[0].strip()
        self.body_length_most_recent_message = len(self.body.split('\r\nFrom:')[0].strip())
        self.conversation_topic = message.ConversationTopic
        self.conversation_id = message.ConversationID
        self.saved = message.Saved
        self.importance = message.Importance
        self.download_state = message.DownloadState
        self.flag_status = message.FlagStatus
        ## Not to be stored
        self.message = message
        if self.attachment_count > 0:
            self.add_attachments()

    def add_attachments(self):
        for attachment in self.message.Attachments:
            idx = attachment.Position
            varname_base = 'attachment{}'.format(str(idx))
            setattr(self, varname_base + '_filename', attachment.FileName)
            setattr(self, varname_base + '_size', attachment.Size)


class Appointment: cals = []


obdict = {}
for idx, x in enumerate(cal):
    if idx < 10:
        obdict['global_appointment_id'] = x.GlobalAppointmentId
        obdict['subject'] = x.Subject
        obdict['location'] = x.Location
        obdict['duration'] = x.Duration
        obdict['busy_status'] = x.BusyStatus
        obdict['attendees_required'] = x.RequiredAttendees
        obdict['attendees_optional'] = x.OptionalAttendees
        obdict['recipients_names'] = [y.name for y in x.Recipients]
        obdict['meeting_status'] = x.MeetingStatus
        obdict['response_status'] = x.ResponseStatus
        obdict['conversation_id'] = x.ConversationID
        obdict['conversation_index'] = x.ConversationIndex
        obdict['conversation_topic'] = x.ConversationTopic
        obdict['start'] = x.Start
        obdict['end'] = x.End
        obdict['unread'] = x.UnRead
        obdict['body'] = x.Body
        obdict['attachment_count'] = x.Attachments.Count
        obdict['all_day_event'] = x.AllDayEvent
        obdict['creation_time'] = x.CreationTime
        obdict['last_modification_time'] = x.LastModificationTime
        obdict['categories'] = x.Categories
        obdict['duration'] = x.Duration
        obdict['is_recurring'] = x.IsRecurring

        cals.append(obdict)
    else:
        break
