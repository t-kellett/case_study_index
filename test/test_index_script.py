import os
import pytest
from ppt_search import find_slide_number, find_slide_number_by_notes

# Define the path to your PowerPoint presentation for testing
TEST_PPTX_PATH = 'test_presentation.pptx'

# Test case for searching slide content
def test_find_slide_number():
    # Replace 'Search Text' with the text you expect to find in the presentation
    slide_number = find_slide_number(TEST_PPTX_PATH, 'Search Text')
    assert slide_number == 2  # Update with the expected slide number

# Test case for searching slide notes
def test_find_slide_number_by_notes():
    # Replace 'Search Text' with the text you expect to find in the notes
    slide_number = find_slide_number_by_notes(TEST_PPTX_PATH, 'Search Text')
    assert slide_number == 3  # Update with the expected slide number

if __name__ == '__main__':
    pytest.main()
