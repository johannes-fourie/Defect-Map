#run this script from .\Defect-Map\Scripts using IronPython
import clr
clr.AddReferenceToFileAndPath('..\\bin\\Release\\TFSBugQuery.dll')
from TFSBugQuery import *

print '("------------ Query Team Foundation Server ------------")'
#QueryBugs.RunQuery([TFS server URL], [TFS project name], [TFS query folder], [TFS query], [lof file path], [optional number of workitems to read from query result, take_first = 0 -> read all items]
QueryBugs.RunQuery('http://tfs:8080/tfs', 'my tfs project', 'My Queries', 'Bugs query', 'C:\\debugger\\Bugs.log')
raw_input("------------ Done ------------")

#ipy RunQuery.py