# Copyright (c) 2010-2020 openpyxl

from openpyxl.compat import safe_string

class DataTable:


    t = "dataTable"

    def __init__(self,
                 ref,
                 dt2D=False,
                 dtr=False,
                 r1=None,
                 r2=None,
                 del1=False,
                 del2=False,
                 **kw):
        self.ref = ref
        self.dt2D = dt2D
        self.dtr = dtr
        self.r1 = r1
        self.r2 = r2
        self.del1 = del1
        self.del2 = del2


    def __iter__(self):
        for k in ["t", "ref", "dt2D", "dtr", "r1", "r2", "del1", "del2"]:
            v = getattr(self, k)
            if v:
                yield k, safe_string(v)
