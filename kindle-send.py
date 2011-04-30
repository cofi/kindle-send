#!/usr/bin/env python2.7
# -*- coding: utf-8 -*-
# (C) 2011 Michael Markert <markert.michael@googlemail.com>
# License: GPL v3 (see http://www.gnu.org/licenses/gpl-3.0.html)
#
# This script allows you to send files -- that are understood by amazon -- to a
# kindle email using msmtp
#
# Used Environment Variables:
# KINDLE_FROM_ADDRESS -- email address of sender
#                        spares using the --from argument
# KINDLE_TO_ADDRESS -- email address of receiving kindle, should end in
#                      @free.kindle.com or @kindle.com
#                      spares using the --to argument
from os.path import splitext, basename
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.Utils import COMMASPACE, formatdate
from email import Encoders
import subprocess

# Map of file extensions amazon supports
# see http://www.amazon.com/gp/help/customer/display.html?nodeId=200505520&#recognize
EXTENSION_TO_MIME_PAIR = {

    # Amazon Kindle
    '.azw'  : ('application', 'octet-stream'),
    '.azw1' : ('application', 'octet-stream'),
    # plain text
    '.txt'  : ('text', 'plain'),
    # mobi
    '.mobi' : ('application', 'octet-stream'),
    '.prc'  : ('application', 'octet-stream'),
    # audible
    '.aa'   : ('application', 'octet-stream'),
    '.aax'  : ('application', 'octet-stream'),
    # MP3
    '.mp3'  : ('application', 'octet-stream'),
    # MS Word
    '.doc'  : ('application', 'msword'),
    '.docx' : ('application', 'msword'),
    # HTML
    '.html' : ('text', 'html'),
    '.htm'  : ('text', 'html'),
    # RTF
    '.rtf'  : ('application', 'rtf'),
    # PDF
    '.pdf'  : ('application', 'pdf'),
    # Images
    '.gif'  : ('image', 'gif'),
    '.png'  : ('image', 'png'),
    '.bmp'  : ('image', 'bmp'),
    '.jpg'  : ('image', 'jpeg'),
    '.jpeg' : ('image', 'jpeg'),
    }

def is_kindle_mail(address):
    return address.endswith("@kindle.com") or address.endswith("@free.kindle.com")

def prepare_mail(attachments):
    """
    Construct a multipart mail that includes all valid attachments.

    :param attachments: List of filenames to be included
    :returns: A mail that includes all (by Amazon supported) attachments.
    :rtype: MIMEMultipart mail with base64-encoded attachments.
    """
    msg = MIMEMultipart()
    filenames = list()
    for attachment in attachments:
        try:
            fname = basename(attachment)
            _, extension = splitext(fname) type_, subtype = EXTENSION_TO_MIME_PAIR[extension]
            part = MIMEBase(type_, subtype)
            with open(attachment, 'rb') as fobj:
                part.set_payload(fobj.read())
            Encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            'attachment; filename="{0}"'.format(fname))
            msg.attach(part)
            filenames.append(fname)
        except KeyError:
            print "Could not prepare {0} for sending. \
'{1}' is not supported by Amazon".format(attachment, extension)

    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = "Transfer {0}".format(', '.join(filenames))
    if filenames:
        return msg
    print "No file was accepted."

def send_mail(from_, to, filenames, account, send=True):
    """
    Send email with files using account.

    :param from_: sending email address.
    :param to: kindle email address.
    :param filenames: List of filenames to be included.
    :param account: msmtp account to use.
    :param send: actually send email.
    """
    msg = prepare_mail(filenames)
    if not is_kindle_mail(to):
        print "Warning: Specified receiver is not a kindle address."
    if not msg or not send:
        print "No mail will be sent."
        return
    msg['From'] = from_
    msg['To'] = to
    print "Sending: {0}.".format(msg['Subject'])
    mailer = subprocess.Popen(['msmtp', '--account={0}'.format(account), to],
                              stdin=subprocess.PIPE)
    mailer.communicate(msg.as_string())

def main():
    import sys
    import os
    import argparse
    parser = argparse.ArgumentParser(description='Send files to kindle email.')
    parser.add_argument('-f', '--from', nargs='?', type=str, dest='from_',
                        default=os.environ.get('KINDLE_FROM_ADDRESS'),
                        help='Sender of email.' + ' ' +
                        'Defaults to $KINDLE_FROM_ADDRESS')
    parser.add_argument('-t', '--to', nargs='?', type=str,
                        default=os.environ.get('KINDLE_TO_ADDRESS'),
                        help='Kindle email address.' + ' ' +
                        'Defaults to $KINDLE_TO_ADDRESS.')
    parser.add_argument('-a', '--account', nargs='?', type=str,
                        default='default',
                        help='MSMTP account used for sending.')
    parser.add_argument('--dry-run', dest='send', action='store_false',
                        help='Send no mail.')
    parser.add_argument('files', metavar="FILE", nargs='+', type=str,
                        help='Files to send')
    args = parser.parse_args()
    failed = False
    if not args.from_:
        print 'No sender specified. Use --from or $KINDLE_FROM_ADDRESS.'
        failed = True
    if not args.to:
        print 'No kindle email address specified. Use --to or $KINDLE_TO_ADDRESS.'
        failed = True
    if failed:
        sys.exit(1)
    send_mail(args.from_, args.to, args.files, args.account, args.send)

if __name__ == '__main__':
    main()
