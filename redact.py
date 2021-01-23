#!/usr/bin/env python

import argparse
import re
import sys

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
from run_tools import *


def process_matches(match_pairs: list, run_text: str):
    """
    Process matchPairs from regex finditer
    Args:
        match_pairs: List of match (start, end) from a regex finditer
        run_text: String

    Returns: Tuple of (boolean list of highlights, number list of positions)
    """
    # List to hold the indices to split at
    matches = []
    # List to hold whether to highlight or not
    # In other words, whether the split contains a pattern
    highlights = []
    if len(match_pairs) > 0:
        # If the first match does not start with zero, append zero
        # highlight as False
        if match_pairs[0][0] != 0:
            matches.append(0)
            highlights.append(False)
        # First match start index and highlight as true
        matches.append(match_pairs[0][0])
        highlights.append(True)
        # First match end index
        matches.append(match_pairs[0][1])
        prev = match_pairs[0][1]
        # Loop through the remaining pairs, except the last
        for idx in range(1, len(match_pairs)):
            # IF the start of the next match does NOT match the end of the previous match
            # THEN some text in the run does not match the pattern and should not be highlighted
            #   AND go to the start of the current match
            # IF the start of the next match does match the end of the previous match
            # THEN the text starting with the start of the match should be highlighted
            if prev != match_pairs[idx][0]:
                highlights.append(False)
                matches.append(match_pairs[idx][0])
            highlights.append(True)
            # Append the end of the current match
            # and go over the same logic with the next match
            matches.append(match_pairs[idx][1])
            prev = match_pairs[idx][1]
        # IF the end of the last match does not match the length of the run-text
        # THEN don't highlight and append the matches with the length of the run-text
        if matches[-1] != len(run_text):
            highlights.append(False)
            matches.append(len(run_text))
    return highlights, matches


def redact_colors(color: str = None):
    """
    Returns text and text-background colors for redaction using argument {color}.
    Args:
        color: str

    Returns: Tuple with redaction colors for (text, background)
    """
    switcher = {
        'white': (RGBColor(255, 255, 255), WD_COLOR_INDEX.WHITE),
        'yellow': (RGBColor(255, 255, 0), WD_COLOR_INDEX.YELLOW)
    }
    return switcher.get(color, (RGBColor(0, 0, 0), WD_COLOR_INDEX.BLACK))


def redact_document(input_path: str, output_path: str, pattern: list,
                    replace_with: str = None, color: str = None):
    """
    Redacts {pattern} after replacing it with {replace_with}
    in the {input} document and saves it as {output} document.
    Args:
        input_path (str): Path to the document to be redacted
        output_path (str): Path to save the redacted document
        pattern (list): List of pattern to redact
        replace_with (str): String to replace all patterns
        color (str): Color to redact. Will be used for both the text and background.
    """

    # Get the text color and text-background color for reaction
    txt_color, background_color = redact_colors(color)

    # Open the input document
    doc = Document(input_path)
    # Loop through paragraphs
    for para in doc.paragraphs:
        # Loop through the runs in the paragraph in the reverse order
        run_index = len(para.runs) - 1
        while run_index > -1:
            run = para.runs[run_index]
            # Find the start and end indices of the patterns in the run-text
            match_pairs = [(match.start(), match.end()) for match in re.finditer('|'.join(pattern), run.text)]
            # Get the locations in the format required for `split_run_by` function
            highlights, matches = process_matches(match_pairs, run.text)
            # Go to redact only if patterns are found in the text
            if len(highlights) > 0 and len(matches) > 0:
                if len(highlights) != len(matches) - 1:
                    ValueError('Calculation error within matches and highlights')
                else:
                    if len(matches) == 2:  # When a pattern is the only text in the run
                        # Replace the matching pattern if substitute given
                        if replace_with is not None:
                            run.text = replace_with
                        # Highlight the background color
                        run.font.highlight_color = background_color
                        # Match the text color to the background color
                        run.font.color.rgb = txt_color
                    else:
                        # Split the runs using the matches
                        new_runs = split_run_by(para, run, matches[1:-1])
                        # Highlight the run if it matches a pattern
                        for highlight, run in zip(highlights, new_runs):
                            if highlight:
                                # Replace the matching pattern if substitute given
                                if replace_with is not None:
                                    run.text = replace_with
                                # Highlight the background color
                                run.font.highlight_color = background_color
                                # Match the text color to the background color
                                run.font.color.rgb = txt_color
            # Decrement the index to process the previous run
            run_index -= 1
    # Save the redacted document to the output path
    doc.save(output_path)


def main(cmd_argv: list):
    """
    Parse command line args and call the right function
    Args:
        cmd_argv (list): List of command line arguments
    """

    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', dest='input',
                        help="Path to the document to be redacted",
                        type=str, required=True)
    parser.add_argument('-o', '--output', dest='output',
                        help="Path to save the redacted document",
                        type=str, required=True)
    parser.add_argument('-p', '--patterns', dest='patterns',
                        help="List of pattern to redact. Comma separated.",
                        type=str, required=True)
    parser.add_argument('-r', '--replace', dest='replace',
                        help="String to replace all patterns",
                        type=str)
    parser.add_argument('-c', '--color', dest='color',
                        help="Color to redact. Black is default. Options: white, yellow",
                        type=str)
    args = parser.parse_args(cmd_argv)

    patterns = [x.strip() for x in args.patterns.split(',')]
    redact_document(args.input, args.output, patterns, args.replace, args.color)


if __name__ == "__main__":
    main(sys.argv[1:])
