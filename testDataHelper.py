import hashlib
import os
import random
import time
import urllib
import string
from datetime import datetime

import SystemConfig

def create_cookie(value):
    pairs = value.split(",")
    cookie = {}
    for pair in pairs:
        [name, value] = pair.split(":")
        cookie[name] = value
    return cookie

def generate_timestamp(timeformat):
    cur_timestamp = datetime.now()

    if "epoch" in timeformat:
        return int(time.time())

    if "YYYY" in timeformat:
        timeformat = timeformat.replace("YYYY", "%Y")

    if "YY" in timeformat:
        timeformat = timeformat.replace("yy", "%y")

    if "yy" in timeformat:
        timeformat = timeformat.replace("yy", "%#y")

    if "MONTH" in timeformat:
        timeformat = timeformat.replace("MONTH", "%B")

    if "MM" in timeformat:
        timeformat = timeformat.replace("MM", "%m")

    if "mm" in timeformat:
        timeformat = timeformat.replace("mm", "%#m")

    if "DD" in timeformat:
        timeformat = timeformat.replace("DD", "%d")

    if "dd" in timeformat:
        timeformat = timeformat.replace("dd", "%#d")

    if "HH" in timeformat:
        timeformat = timeformat.replace("HH", "%H")

    if "HM" in timeformat:
        timeformat = timeformat.replace("HM", "%I")

    if "MI" in timeformat:
        timeformat = timeformat.replace("MI", "%M")

    if "mi" in timeformat:
        timeformat = timeformat.replace("mi", "%#M")

    if "SS" in timeformat:
        timeformat = timeformat.replace("SS", "%S")

    if "ss" in timeformat:
        timeformat = timeformat.replace("ss", "%#S")

    return cur_timestamp.strftime(timeformat)

def encode_string(text, encode_type):
    encoded_string = ""
    if encode_type == "sha1":
        key = hashlib.sha1()
        key.update(text)
        encoded_string = key.hexdigest()
    else:
        print "[WARN] {0} is not yet supported".format(encode_type)
    return encoded_string

def split(text, delimiter, index):
    text = text.split(delimiter)
    return text[int(index)]

def theia_double_encode():
    os.system("getEncodedVal.pys")

    urllib.quote_plus('W7Bv+KOF0xQIgf2T2V/LJQ==')

def random_int(min_value, max_value):
    return random.randint(int(min_value), int(max_value))

def random_value(prefix, number_of_chars, suffix, exclusions, pool):
    random_pool = ""
    if pool == "":
        random_pool = "string.digits"

    if "numbers" in pool:
        random_pool += string.digits
    if "alpha" in pool:
        random_pool += string.ascii_letters
    if "special" in pool:
        random_pool += string.punctuation

    for char in SystemConfig.fixedExclusions:
        random_pool = random_pool.replace(char, "")
    for char in exclusions:
        random_pool = random_pool.replace(char, "")

    number_of_chars = int(number_of_chars) - len(prefix) - len(suffix)

    filler = ""
    if number_of_chars > 0:
        itr = 0
        while itr < number_of_chars:
            filler += random.choice(random_pool)
            itr += 1
    return prefix + filler + suffix