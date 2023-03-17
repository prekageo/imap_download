#!/usr/bin/env python3

import email
import email.utils
import getpass
import hashlib
import imaplib
import os
import re
import sqlite3
import sys
import time

import tqdm


# Copied from https://github.com/OfflineIMAP/offlineimap3
def __split_quoted(s):
    """Looks for the ending quote character in the string that starts
    with quote character, splitting out quoted component and the
    rest of the string (without possible space between these two
    parts.

    First character of the string is taken to be quote character.

    Examples:
     - "this is \" a test" (\\None) => ("this is \" a test", (\\None))
     - "\\" => ("\\", )
    """

    if len(s) == 0:
        return b"", b""

    q = quoted = bytes([s[0]])
    rest = s[1:]
    while True:
        next_q = rest.find(q)
        if next_q == -1:
            raise ValueError("can't find ending quote '%s' in '%s'" % (q, s))
        # If quote is preceeded by even number of backslashes,
        # then it is the ending quote, otherwise the quote
        # character is escaped by backslash, so we should
        # continue our search.
        is_escaped = False
        i = next_q - 1
        while i >= 0 and rest[i] == ord("\\"):
            i -= 1
            is_escaped = not is_escaped
        quoted += rest[0 : next_q + 1]
        rest = rest[next_q + 1 :]
        if not is_escaped:
            return quoted, rest.lstrip()


# Copied from https://github.com/OfflineIMAP/offlineimap3
def imapsplit(imapstring):
    """Takes a string from an IMAP conversation and returns a list containing
    its components.  One example string is:

    (\\HasNoChildren) "." "INBOX.Sent"

    The result from parsing this will be:

    ['(\\HasNoChildren)', '"."', '"INBOX.Sent"']"""

    workstr = imapstring.strip()
    retval = []
    while len(workstr):
        # handle parenthized fragments (...()...)
        if workstr[0] == ord("("):
            rparenc = 1  # count of right parenthesis to match
            rpareni = 1  # position to examine
            while rparenc:  # Find the end of the group.
                if workstr[rpareni] == ord(")"):  # end of a group
                    rparenc -= 1
                elif workstr[rpareni] == ord("("):  # start of a group
                    rparenc += 1
                rpareni += 1  # Move to next character.
            parenlist = workstr[0:rpareni]
            workstr = workstr[rpareni:].lstrip()
            retval.append(parenlist)
        elif workstr[0] == ord('"'):
            # quoted fragments '"...\"..."'
            (quoted, rest) = __split_quoted(workstr)
            retval.append(quoted)
            workstr = rest
        else:
            splits = workstr.split(maxsplit=1)
            splitslen = len(splits)
            # The unquoted word is splits[0]; the remainder is splits[1]
            if splitslen == 2:
                # There's an unquoted word, and more string follows.
                retval.append(splits[0])
                workstr = splits[1]  # split will have already lstripped it
                continue
            elif splitslen == 1:
                # We got a last unquoted word, but nothing else
                retval.append(splits[0])
                # Nothing remains.  workstr would be ''
                break
            elif splitslen == 0:
                # There was not even an unquoted word.
                break
    return retval


# Copied from https://github.com/OfflineIMAP/offlineimap3
def dequote(s):
    """Takes string which may or may not be quoted and unquotes it.

    It only considers double quotes. This function does NOT consider
    parenthised lists to be quoted."""

    if s and s.startswith(b'"') and s.endswith(b'"'):
        s = s[1:-1]  # Strip off the surrounding quotes.
        s = s.replace(b'\\"', b'"')
        s = s.replace(b"\\\\", b"\\")
    return s


# Same as imaplib.Internaldate2tuple without the parsing time zone.
def Internaldate2tuple(resp):
    InternalDate = re.compile(
        rb'.*INTERNALDATE "'
        rb"(?P<day>[ 0123][0-9])-(?P<mon>[A-Z][a-z][a-z])-(?P<year>[0-9][0-9][0-9][0-9])"
        rb" (?P<hour>[0-9][0-9]):(?P<min>[0-9][0-9]):(?P<sec>[0-9][0-9])"
        rb" \+0000"
        rb'"'
    )

    Months = " Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(" ")
    Mon2num = {s.encode(): n + 1 for n, s in enumerate(Months[1:])}

    mo = InternalDate.match(resp)
    if not mo:
        return None

    day = int(mo.group("day"))
    mon = Mon2num[mo.group("mon")]
    year = int(mo.group("year"))
    hour = int(mo.group("hour"))
    min = int(mo.group("min"))
    sec = int(mo.group("sec"))

    return (year, mon, day, hour, min, sec, 0, 1, -1, 0)


class LocalFolder:
    def __init__(self, path):
        self.path = path
        if not os.path.exists(self.path):
            os.makedirs(self.path)

    def count(self):
        return len(os.listdir(self.path))

    def get_existing(self):
        filenames = os.listdir(self.path)
        return [int(f.split(".")[0]) for f in filenames]

    def store(self, uid, data, metadata):
        filename = os.path.join(self.path, f"{uid}.eml")
        assert not os.path.exists(filename)
        with open(filename, "wb") as f:
            f.write(data)
        t = 0
        e = email.message_from_bytes(data)
        date = e["Date"]
        time_tuple = None
        if date is not None:
            time_tuple = email.utils.parsedate_tz(str(date))
        if time_tuple is None:
            print(f"{filename} has no Date, will use INTERNAL_DATE", file=sys.stderr)
            time_tuple = Internaldate2tuple(metadata)

        t = email.utils.mktime_tz(time_tuple)
        os.utime(filename, (t, t))


def connect(host, user, password):
    imap = imaplib.IMAP4_SSL(host)
    resp = imap.login(user, password)
    assert resp[0] == "OK"
    return imap


def get_folders(imap):
    folders = []
    resp = imap.list()
    assert resp[0] == "OK"
    for item in resp[1]:
        flags, delim, name = imapsplit(item)
        folders.append(name)
    return folders


def download(imap, destination, conn):
    folders = get_folders(imap)

    for folder in folders:
        folder_unquoted = dequote(folder).decode()
        if folder_unquoted in ["Bulk Mail", "Bulk"]:
            continue
        resp = imap.select(folder, readonly=True)
        assert resp[0] == "OK"
        count = int(resp[1][0])

        local_folder = LocalFolder(os.path.join(destination, folder_unquoted))
        existing_count = local_folder.count()

        if count != existing_count:
            print(
                f"{folder_unquoted} missing {count - existing_count}", file=sys.stderr
            )

            resp = imap.uid("search", None, "ALL")
            assert resp[0] == "OK"
            remote_uids = set(map(int, resp[1][0].split()))

            local_uids = set(local_folder.get_existing())
            assert local_uids.issubset(remote_uids)

            missing_uids = sorted(remote_uids - local_uids)
            assert len(missing_uids) == count - existing_count

            for msg_uid in tqdm.tqdm(missing_uids):
                resp = imap.uid(
                    "FETCH",
                    str(msg_uid),
                    "(emailid threadid FLAGS INTERNALDATE BODY.PEEK[])",
                )
                assert resp[0] == "OK"
                metadata, data = resp[1][0]
                local_folder.store(msg_uid, data, metadata)
                args = (
                    folder_unquoted,
                    msg_uid,
                    metadata.decode(),
                    hashlib.sha1(data).hexdigest(),
                )
                cur = conn.cursor()
                cur.execute(
                    'insert into emails(folder, uid, metadata, created_at, sha1) values (?,?,?,datetime("now"),?)',
                    args,
                )
                conn.commit()
                time.sleep(1)


def main():
    host = ""
    user = ""
    password = getpass.getpass()
    destination = ""

    conn = sqlite3.connect("storage.sqlite")
    conn.row_factory = sqlite3.Row
    conn.execute(
        "create table if not exists emails(folder, uid, metadata, created_at, sha1)"
    )

    while True:
        try:
            imap = connect(host, user, password)
            download(imap, destination, conn)
            break
        except imaplib.IMAP4.abort as e:
            print(e)
            time.sleep(60)


if __name__ == "__main__":
    main()
