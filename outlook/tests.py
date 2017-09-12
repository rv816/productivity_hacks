from .src.extract import inbox, cal, sent
from .src.transform import Message, Appointment


def test_basic():
    thing = inbox  # import successful
    assert True
# Create your tests here.

def test_inbox():
    first = next(enumerate(inbox))  # wraps inbox into iterator
    m = Message(first[1])
    assert type(m.body) == str


def test_cal():
    first = next(enumerate(cal))  # wraps inbox into iterator
    a = Appointment(first[1])
    assert type(a.body) == str


def test_sent():
    first = next(enumerate(sent))  # wraps inbox into iterator
    m = Message(first[1])
    assert type(m.body) == str
