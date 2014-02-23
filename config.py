class Config:
    Username = None
    Password = None
    FromDate = None
    ToDate = None
    Label = None
    Orders = []
    Total_Amount = None
    Revenue_Amount = None

import sys
import os
import imp


def main_is_frozen():
    return (hasattr(sys, "frozen") or  # new py2exe
            hasattr(sys, "importers")  # old py2exe
            or imp.is_frozen("__main__"))  # tools/freeze


def get_main_dir():
    if main_is_frozen():
        return os.path.dirname(os.path.dirname(sys.executable))
    return os.path.abspath(os.path.dirname(sys.argv[0]))