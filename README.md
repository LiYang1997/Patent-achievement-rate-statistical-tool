1.安装环境：
	需要安装python环境 并下载openpyxl库，因为excel操作依赖这个库

2.使用说明：
	执行专利达成率统计.py 后面加参数
	 -c代表创建文件，后跟文件名xxx.xlsx 
	 -team-objectives代表团队目标 后面跟总共需要完成的专利提案份数
	 -name代表团队所有成员名字 后面跟名字列表
	 例如：专利达成率统计.py -c test.xlsx -team-objectives 10 -names 李洋 荀玉朝 武小虎 韩青山 李鹏博
	 	会生成所有人对应的一个表格，大家达成率都是0.00%
	-f代表要指定操作的文件名字 后面跟文件名xxx.xlsx
	-idea代表谁想出的idea以及份数 后面跟表格中已有的名字以及要完成的份数
	例如：专利达成率统计.py -f test.xlsx -idea 李洋 1
		会将李洋从旧的完成率列表移动到新的完成率列表，表格最左边一列到最右边一列的顺序为从完成率最高人员到完成率最低人员排序
3.待办：
	目前只实现了创建文件和统计提出idea的功能，未实现增加人数，删除人数，idea无效删除idea等功能，这些暂时手动调整