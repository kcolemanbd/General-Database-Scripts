# ADODB-Connection-in-Python
# ADODB Connection written in Python

from System import *
from System.Windows.Forms import * #Used for message boxes during testing

class DataBaseConnect(object):
	def get_RSCount(self):
		return self._RecordCount

	RSCount = property(fget=get_RSCount)

	def get_AbsPosition(self):
		return self._AbsolutePosition

	AbsPosition = property(fget=get_AbsPosition)

	def get_EOF(self):
		return self._rs.EOF

	EOF = property(fget=get_EOF)

	def get_BOF(self):
		return self._rs.BOF

	BOF = property(fget=get_BOF)

	def get_Bookmark(self):
		return self._BMark

	def set_Bookmark(self, value):
		self._BMark = value

	Bookmark = property(fget=get_Bookmark, fset=set_Bookmark)

	def __init__(self):
		self._cn = ADODB.ConnectionClass()
		self._rs = ADODB.RecordsetClass()
		self._RecordCount = 0
		self._AbsolutePosition = 0
		self._BMark = 0



	def Connect(DBName):
		try:
			DBProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
			DBSource = "Data Source=" + Application.StartupPath.ToString() + DBName + ";"
			ConStr = DBProvider + DBSource
			self._cn.Open(ConStr, "", "", 0)
		except Exception, e:
			MessageBox.Show(e.ToString())
		finally:

	Connect = staticmethod(Connect)






def RSConnect(self, TableName):
		try:
			rs.Open(TableName, cn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, 0)
			if rs.EOF == True and rs.BOF == True:
				RecordCount = 0
				AbsolutePosition = 0
			else:
				rs.MoveFirst()
				while rs.EOF == False:
					RecordCount += 1
					rs.MoveNext()
				rs.MoveFirst()
		except , :
			MessageBox.Show("Cannot connect to table specified.")
		finally:




	def Value(self, FieldNum):
		return (rs.Fields[FieldNum].Value.ToString())




	def Navigate(self, Direction):
		if Direction.ToUpper() == "FIRST":
			rs.MoveFirst()
			AbsolutePosition = 1
		elif Direction.ToUpper() == "PREVIOUS" or Direction.ToUpper() == "PREV":
			if rs.BOF == False:
				rs.MovePrevious()
				AbsolutePosition -= 1
		elif Direction.ToUpper() == "NEXT":
			if rs.EOF == False:
				rs.MoveNext()
				AbsolutePosition += 1
		elif Direction.ToUpper() == "LAST":
			rs.MoveLast()
			AbsolutePosition = RecordCount
		else:
			MessageBox.Show(Direction + " is not a valid argument.")
 
 
 
def GotoBookmark(self):
		rs.MoveFirst()
		AbsolutePosition = 1
		i = 1
		while i < BMark:
			rs.MoveNext()
			AbsolutePosition += 1
			i += 1
