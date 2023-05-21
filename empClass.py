class Emp:
    def __init__(self, eid, enm, eage, esal, eadd):
        self.EmpId = eid
        self.EmpName = enm
        self.EmpAge = eage
        self.EmpSalary = esal
        self.EmpAddress = eadd

    def __str__(self):
        return f'{self.EmpId, self.EmpName, self.EmpAge, self.EmpSalary, self.EmpAddress}'

    def __repr__(self):
        return str(self)
