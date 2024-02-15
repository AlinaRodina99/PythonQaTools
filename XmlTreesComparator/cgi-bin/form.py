#!/usr/bin/env python3
import cgi


def ohash(s):
    ans = 0
    for c in s:
        ans = ans * 123417 + ord(c)
    return ans


form = cgi.FieldStorage()
name = form.getfirst("Input text", "Не задано")

print("Content-type: text/html\n")
print("""<!DOCTYPE HTML>
        <html>
        <head>
            <meta charset="utf-8">
        </head>
        <body>""")

print("<h1>" + str(ohash(name)) + "</h1>")

print("""</body>
        </html>""")