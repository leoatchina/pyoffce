    # -*- coding: utf-8 -*-
# @Author: leoatchina
# @Date:   2016-06-14 21:28:24
# @Last Modified by:   leoatchina
# @Last Modified time: 2016-08-19 12:36:51
__author__ = 'leoatchina'


def kwargs(**kwargs):
    return kwargs

from tools import *

class MDB():
    def __init__(self,
                    con=None,
                    host='127.0.0.1',
                    user='test',
                    passwd='test',
                    db='test',
                    port=3306,
                    charset='utf8'
                ):

        import MySQLdb
        if con is None:
            if not isinstance(port,int):
                try:
                    port =int(port)
                except:
                    raise Exception("port is not integer")
            self.con=MySQLdb.connect(host=host,
                                    user=user,
                                    passwd=passwd,
                                    db=db,
                                    port=port,
                                    #估计下面这个charset就是原来出错的地方
                                    charset=charset)
        else:
            self.con=con
        if isinstance(self.con,MySQLdb.connections.Connection):
            self.cur=self.con.cursor()
        else:
            print "init failed"
            self.con=None
            exit(0)

    def execute(self,*args,**kwargs):
        self.cur.execute(*args,**kwargs)
        self.con.commit()
        returnTuple=self.cur.fetchall()
        if returnTuple:
            if len(returnTuple[0])==1:
                returnList= [each[0].encode('utf-8') if isinstance(each[0],unicode) else each[0] for each in returnTuple]
            else:
                returnList= [[i.encode('utf-8') if isinstance(i,unicode) else i for i in each ] for each in returnTuple]
            return returnList
        else:
            return None


    def close(self):
        if self.cur :
            self.cur.close()
            self.con.close()

    def __del__(self):
        if self.cur :
            self.cur.close()
            self.con.close()
    #检查是否在table里，成功的话返回整个表的全部column的list,否则返回空
    def __checkInTable(self,table,*columns,**values):
         #非字符table或者 table里有空格
        if not isinstance(table,str) or r" " in table.strip():
            raise Exception("table %s name format wrong"%table)
        sql=r"show tables like '%s';"%table
        self.cur.execute(sql)
        if  not len(self.cur.fetchall()):
            print r"table '%s' does not exist"%table
        self.cur.execute("SHOW columns FROM %s"%table)
        fetchall=self.cur.fetchall()
        fetchlist=[i[0].encode("utf-8")  for i in fetchall]
        if columns:
            if r"*" in columns:
                pass
            else:
                for column in columns:
                    if not column in fetchlist:
                        raise Exception(r"the table '%s' does not contain the columns '%s'"%(table,column))
        if values:
            for value in values:
                if not value in fetchlist:
                    raise Exception(r"the table '%s' does not contain the columns '%s'"%(table,value))

        return fetchlist




    def insert(self,table,**values):
        return self.__insert(table,True,**values)


    def insertNoCheck(self,table,**values):
        return self.__insert(table,False,**values)


    def __insert(self,table,checkExsists,**values):#insert只会插入一个值
        if not isinstance(values,dict):
            print "insert values is not dict"
            return False

        if checkExsists:
            idList=self.selectOne(table,**values)
            if idList:
                return  idList



        indexList=getDictIndexList(values)
        valueList=getDictValueList(values)
        columnStr=listToStr(indexList)
        valueStr=listToCharStr(valueList)
        insertSql="INSERT INTO %s(%s) Values(%s)"%(table,columnStr,valueStr)
        self.cur.execute(insertSql,valueList)
        self.con.commit()
        # ID=self.cur.execute("SELECT LAST_INSERT_ID();") # 不知道为什么，这个老是返回1
        return self.cur.lastrowid

    def delete(self,table,**wheres):
        if not isinstance(wheres,dict):
            print "delete values is not dict"
            return False
        indexList=self.select(table,**wheres)
        if indexList:# 用sql语句删除
            deleteSql="delete from %s"%table
            valueList=getDictValueList(wheres)
            whereCharStr=" WHERE "+dictToCharStr(wheres)
            self.cur.execute(deleteSql+whereCharStr,valueList)
            self.con.commit()
            return True
        else:
            print "nothing to delete in table:",table
            return False

    #updata有insert的功能
    def update(self,table,wheres,values=None):
        columnList=self.__checkInTable(table,**wheres)
        if not columnList:
            raise Exception("update condition columns wrong")
            return False
        if values:
            columnList=self.__checkInTable(table,**values)
            if not columnList:
                raise Exception("update target columns wrong")
                return False
        idList=self.select(table,**wheres)
        if idList:#说明可以update
            MainID=columnList[0]
            if values:
                valueList=getDictValueList(values)
                valueCharStr=dictToCharStr(values)
                whereList=["%s=%s"%(MainID,i) for i in idList]
                whereStr="WHERE "+listToStr(whereList,False," OR ")
                updateSql =r"UPDATE %s set %s " %(table,valueCharStr)
                self.cur.execute(updateSql+whereStr,valueList)
                self.con.commit()
            else:#成了插入
                self.insert(table,**wheres)
            return True
        else:# update 变成insert，直接把values里的值导入到wheres里去后插入
            if values:
                for i in values:
                    wheres[i]=values[i]
            self.insert(table,**wheres)
            return True


    def count(self,table,**wheres):
        countSql=r"SELECT COUNT(*) FROM %s"%table
        if wheres:
            whereStr=" WHERE "+ dictToCharStr(wheres)
            valueList=getDictValueList(wheres)
            self.cur.execute(countSql+whereStr,valueList)
        else:
            self.cur.execute(countSql)
        return self.cur.fetchone()[0]


    def selectOne(self,table,*columns,**wheres):
        selected=self.select(table,*columns,**wheres)
        if selected:
            return selected[0]
        else:
            return None

    def select(self,table,*columns,**wheres):
        return self.__selectOrigin(table,True,*columns,**wheres)

    def selectNoDistinct(self,table,*columns,**wheres):
        return self.__selectOrigin(table,False,*columns,**wheres)

    def __selectOrigin(self,table,dist,*columns,**wheres):
        #culumnList通过检验后返回的这个表里所有的列的list
        columnList=self.__checkInTable(table,*columns,**wheres)
        if columnList:
            if r"*" in columns:
                length=len(columnList)
            elif not columns:#如果columns为空，那就返回第一列，也就是ID
                columnList=[columnList[0]]
                length=1
            else:
                columnList=[i for i in columns]
                length=len(columnList)
        else:
            return None
        columnStr=r','.join(columnList)
        if dist:
            selectSql="SELECT DISTINCT %s FROM %s"%(columnStr,table)
        else:
            selectSql="SELECT %s FROM %s"%(columnStr,table)
        if wheres:
            whereStr=" WHERE "+ dictToCharStr(wheres)
            valueList=getDictValueList(wheres)
            self.cur.execute(selectSql+whereStr,valueList)
        else:
            self.cur.execute(selectSql)
        returnTuple=self.cur.fetchall()
        if returnTuple:
            if length==1:
                returnList= [each[0].encode('utf-8') if isinstance(each[0],unicode) else each[0] for each in returnTuple]
            else:
                returnList= [[i.encode('utf-8') if isinstance(i,unicode) else i for i in each ] for each in returnTuple]
            return returnList
        else:
            return None
