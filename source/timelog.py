"""
Created on Aug 2, 2016

@usage:
starttime = timelog.now() -> start of block of code
timelog.adhoclog("block of code name", starttime) -> end of block of code
"""

import atexit
from time import time
from datetime import timedelta


def seconds_to_str(t):
    return str(timedelta(seconds=t))


def log(msg, elapsed=None):
    line = "=" * 40
    print(line)
    print(seconds_to_str(time()), '-', msg)
    if elapsed:
        print("Elapsed time:", elapsed)
    print(line)
    print()


def adhoclog(msg, starttime):
    elapsed = time() - starttime
    log(msg, seconds_to_str(elapsed))


def endlog():
    end = time()
    elapsed = end-start
    log("End Program", seconds_to_str(elapsed))


def now():
    # return seconds_to_str(time())
    return time()

start = time()
atexit.register(endlog)
log("Start Program")

