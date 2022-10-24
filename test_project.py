import project
import pytest
import datetime
import os
from unittest.mock import MagicMock

FAKE_NOW = datetime.datetime(2022, 10, 21, 22, 22, 22)

"""Mock object for the fake current time"""
@pytest.fixture()
def mock_print_date(monkeypatch):
    print_date_mock = MagicMock(wraps = datetime.datetime)
    print_date_mock.now.return_value = FAKE_NOW
    monkeypatch.setattr(datetime, "datetime", print_date_mock)

def test_get_print_date(mock_print_date):
    """Checks the programs returns the current date"""
    assert project.get_print_date() == FAKE_NOW

def test_amount_in_words():
    """"Checks if the program returns the correct amount in words using inflect"""
    assert project.amount_in_words("2000") == "TWO THOUSAND"
    assert project.amount_in_words("9750") == "NINE THOUSAND, SEVEN HUNDRED AND FIFTY"

def test_merge_pdfs():
    """Checks if the directory and the file for merged PDFs exists"""
    assert os.path.isdir("D:\\Python Projects\\project\\merged_dir") == True
    assert os.path.isfile("D:\\Python Projects\\project\\merged_dir\\merged_pdfs.pdf") == True

def test_create_pdf():
    """Checks if the output directory and individual PDFs exists"""
    assert os.path.isdir("D:\\Python Projects\\project\\output") == True
    assert os.listdir("D:\\Python Projects\\project\\output") != None
