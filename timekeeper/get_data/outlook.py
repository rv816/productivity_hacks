import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # "6" refers to the index of a folder - in this case,
# the inbox.
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


# test
class Appointment:
    def __init__(self, x):
        """ 
        :param x:"Appointment Item" 
        """
        self.global_appointment_id = x.GlobalAppointmentId
        self.subject = x.Subject
        self.location = x.Location
        self.duration = x.Duration
        self.busy_status = x.BusyStatus
        self.attendees_required = x.RequiredAttendees
        self.attendees_optional = x.OptionalAttendees
        self.recipients_names = [y.name for y in x.Recipients]
        self.meeting_status = x.MeetingStatus
        self.response_status = x.ResponseStatus
        self.conversation_id = x.ConversationID
        self.conversation_index = x.ConversationIndex
        self.conversation_topic = x.ConversationTopic
        self.start = x.Start
        self.end = x.End
        self.unread = x.UnRead
        self.body = x.Body
        self.attachment_count = x.Attachments.Count
        self.all_day_event = x.AllDayEvent
        self.creation_time = x.CreationTime
        self.last_modification_time = x.LastModificationTime
        self.categories = x.Categories
        self.duration = x.Duration
        self.is_recurring = x.IsRecurring
