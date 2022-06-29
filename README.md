# Zhejiang_China_Gaokao_Rank_Copy
Codes to copy the previous years rankings of majors into current year's admission plans based on same major name. 

这是一份添加好了所有2021 年同学校同专业排名（除去新的专业、新的学校名字）的高考 2022 招生计划。可以用两种方法对这张表进行测漏：

1. 打开"浙江省历年普通类第一段平行投档分数线表.xlsx"，检查"Ranking_Copy_Output.xlsx"中的专业排名是否正确，检查标记为“新专业名”排名是否的确是 2021 年没有出现过的；
2. 打开“2022年_普通类平行计划1_普通类平行录取_地理生物技术.xlsm"，检查"Ranking_Copy_Output.xlsx"所有的大学与专业是否都正确被录入；

整个代码记录了7674 条专业的 2021 年排名，占比62.4%。其中排名为999999的专业是因为这个专业在去年掉入了二段，所以默认给了最大排名值。
