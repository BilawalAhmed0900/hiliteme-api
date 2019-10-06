#!/usr/bin/env python3
import argparse
import requests

try:
    import win32clipboard
    import win32con
    FOUND_WIN32_CLIPBOARD = True
except ImportError:
    FOUND_WIN32_CLIPBOARD = False


def main():
    parser = argparse.ArgumentParser(allow_abbrev=False)
    
    input = parser.add_mutually_exclusive_group(required=True)
    output = parser.add_mutually_exclusive_group(required=True)
    
    input.add_argument("--File", metavar="File", help="File to format")
    output.add_argument("--OutputFile", metavar="OutputFile", help="HTML file to output")
    
    parser.add_argument("--lexer", help="Lexer to be used(see lexer.txt for options, default python)",
                        default="python")
    parser.add_argument("--style", help="Color style to use(see style.txt for options, default vs)",
                        default="vs")
    parser.add_argument("--linenos", help="Include line number", action="store_true")
    if FOUND_WIN32_CLIPBOARD:
        output.add_argument("--copytoclipboard", help="Copy to clipboard(windows, pywin32 only)", action="store_true")
        input.add_argument("--copyfromclipboard", help="Copy from clipboard(windows, pywin32 only)", action="store_true")

    arg = parser.parse_args()

    if arg.copyfromclipboard and FOUND_WIN32_CLIPBOARD:
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT).encode()
        win32clipboard.CloseClipboard()
    else:
        file_str = arg.File
        file_obj = open(file_str, "rb")
        text = file_obj.read()
        file_obj.close()

    data = {"code": text.decode(), "lexer": arg.lexer, "style": arg.style, "divstyles": ""}
    if arg.linenos:
        data["linenos"] = "true"

    r = requests.post("http://hilite.me/api", data=data)
    text = r.content

    if arg.copytoclipboard and FOUND_WIN32_CLIPBOARD:
        clipboard_data_format = r"Version:0.9\r\n" \
                                 "StartHTML:{:09d}\r\n" \
                                 "EndHTML:{:09d}\r\n" \
                                 "StartFragment:{:09d}\r\n" \
                                 "EndFragment:{:09d}\r\n" \
                                 "<!doctype>\r\n" \
                                 "<html>\r\n" \
                                 "<body>\r\n" \
                                 "<!--StartFragment -->\r\n" \
                                 "{}\r\n" \
                                 "<!--EndFragment -->\r\n" \
                                 "</body>\r\n" \
                                 "</html>\r\n"

        prefix = r"Version:0.9\r\n" \
                  "StartHTML:{:09d}\r\n" \
                  "EndHTML:{:09d}\r\n" \
                  "StartFragment:{:09d}\r\n" \
                  "EndFragment:{:09d}\r\n"

        text = text.decode()

        data_for_prefix = prefix.format(0, 0, 0, 0)
        len_prefix = len(data_for_prefix)

        data_for_prefix = clipboard_data_format.format(0, 0, 0, 0, text)
        data_len = len(text)

        start_data = data_for_prefix.index(text)
        end_data = start_data + data_len
        actual_formatted_data = clipboard_data_format.format(len_prefix, len(data_for_prefix),
                                                             start_data, end_data,
                                                             text)

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"), actual_formatted_data.encode("UTF-8"))
        win32clipboard.CloseClipboard()
    else:
        file_obj = open(arg.OutputFile, "wb")
        file_obj.write(text)
        file_obj.close()


if __name__ == "__main__":
    main()
