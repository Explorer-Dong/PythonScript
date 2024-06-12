# 基于python的openpyxl库实现志愿审核自动化

## 前言

在介绍脚本编写背景之前，我们先来回顾一下各学院的志协工作人员进行志愿审核的大致步骤，一般有两种：

第一种，纯手动：在”志愿活动审核“一栏中的”待院团委审核“一栏中，对照申请的同学 `applicant` 与实际名单中的同学 `students` 进行对照检索，如果 `applicant` 在 `students` 中，则将该申请者通过院团委审核，然后进入已通过的名单中找到刚才审核通过的同学赋予 `students` 中 `applicant.duration` 的志愿时长。**适用于**小规模活动或者补录、特殊活动等。

第二种，半自动：我们首先获取三个 `excel` 文件

1. 获取申请表 `applicants.xls`。在”志愿活动审核“一栏中的”待院团委审核“一栏中，输入待审核活动的活动名称，将过滤出来的信息全部通过，然后进入“已通过”界面，点击导出，默认导出文件名为 `zyhdshcx_display.xls`，将其重命名为 `applicants.xls`，表示申请的同学名录。
2. 获取导入模板 `template.xls`。在“志愿活动发起”一栏中，任选一个活动，点击“名单”，然后点击“更新导入”，下载导入模板，默认文件名为 `ZYHD_MD_DR.xls`，我们将其重命名为 `template.xls`，表示导入模板。
3. 获取活动名录表 `<ActivityName>.xls` or `<ActivityName>.xlsx`。直接从腾讯文档中下载相关的活动名录文件即可。

接下来我们需要逐个对照每一位申请的同学 `applicant` 与是否存在于 `<ActivityName>` 中，如果存在就将该同学的学号、姓名和志愿时长复制到 `template.xls` 中的相关字段之下，如果不存在就跳过这个人。最后将 `template.xls` 通过“志愿活动发起”一栏找到相关活动直接导入即可完成相应的志愿时长的赋值。**适用于**大规模活动。

~~可能会有同学好奇，为什么还需要同学们申请才能被审核从而获取志愿时长，而非直接将做志愿的同学赋予志愿时长，跳过自己申请这一环节呢（因为很容易忘记申请啊！）？个人认为这是学校一站式网站志愿审核板块开发逻辑的一个漏洞，明明可以直接跳过个人申请这一环节的，利用学号的唯一性直接进行映射赋值即可，很明显这是一个很简单的实现。~~

## 背景

脚本的开发背景是基于上述第二种半自动化展开的。可以发现，其中最繁琐且耗费精力的是“检查每一位 `applicant` 与是否存在于 `<ActivityName>` 中并复制粘贴到 `template.xls` 中的相关字段之下”的流程。尤其是人数过多以后，即使是 Ctrl+F 也颇费时间与精力。因此基于此现象，我将利用 python 的 openpyxl 第三方库自动化完成以上步骤

## 使用逻辑

创建 python 虚拟环境，建立以下 ==TODO== 文件组织架构

首先**手动**下载 3 个文件。按照上述第二种半自动审核的方式将三个 excel 文件全部下载到项目的 files 文件夹下，并进行重命名，分别为 `applicants.xls`、`template.xls` 和 `<ActivityName>.xls`。

接下来**自动化**对照。对照每一个 `applicant` 是否存在于 `<ActivityName>.xls` 中，直接按照 `applicants.xls` 中每一个 `applicant` 的学号与姓名在 `<ActivityName>.xls` 中学号姓名两个字段下进行查询遍历，如果学号或者姓名匹配到了，则将这位 `applicant` 在 `<ActivityName>.xls` 的学号、姓名、活动名称和志愿时长复制到 `template.xls` 中；如果学号或者姓名都没有匹配上，则将这位 applicant 在 `<ActivityName>.xls` 的学号、姓名、活动名称和志愿时长复制到 `template.xls` 中，其中志愿时长赋为 0，并将这条信息在 `<ActivityName>.xls` 标为黄色。（匹配时使用了学号与姓名可以规避掉两个问题，如果只按照学号进行匹配，可能会有前导零的问题；如果只按照姓名进行匹配，可能会有重名现象导致检索出错，综合二者可以很大程度上规避上述问题。匹配失败后在 `<ActivityName>.xls` 中标黄可以后期人工检查程序的正确性，以及警示相关申请错误的同学不要胡乱申请）。

最后**手动**格式化并上传文件。由于一站式中对于上传的模板文件有严格的格式要求，因此我们需要格式化一下 `template.xls` 文件。即将学号、姓名两栏调整为文本格式，然后更新导入到一站式即可。

提示。当前开发逻辑是以单个活动为基元进行，无法进行批量活动的检查审核，因此每次都需要重新下载相关活动的申请表 `applicants.xls` 与 活动名录表 `<ActivityName>.xls` or `<ActivityName>.xlsx`。至于导入模板 `template.xls` 文件，第一次导入后就不需要再到网上下载了，脚本在每次运行之前都会删除 `template.xls` 中除了字段以外的所有信息。

## 代码逻辑

封装到 Process 类中，变量为 3 个 dataframe 用来存储三个文件表，applicants、activity 和 template，类方法如下：

- 文件类型转换方法。由于有些文件是 .xls 后缀，有些是 .xlsx 后缀，为了统一以及适配最新的第三方包，统一将 .xls 后缀文件转化为 .xlsx 文件，经过测试，一站式上传 .xlsx 模板文件也是可以的。编写 xls2xlsx 函数。

- 读取相关信息方法。读取 applicants 中的学号与姓名 2 列信息 apply_name, apply_id，读取 activity 中的学号、姓名和志愿时长 3 列信息 activity_name, activity_id, activity_duration。读取活动名称 activity_title 为 `<ActivityName>.xlsx` 的文件名。编写 read 函数。
- 写入模板信息方法。遍历 apply_name 与 apply_id 列表的 name 与 id，如果 name in activity_name or id in activity_id，则将 `name` `id` `activity_name` `activity_duration` 加入到 `template.xlsx` 文件相关字段的后面即可；反之同理，但是有两点不同。其一为，需要将  `activity_duration` 赋为 0；其二为，需要在 `<ActivityName>.xlsx` 文件中将这条信息填充为黄色背景。编写 write 函数。

## 后续展望

- 支持批量活动的文件操作
- 支持图形化操作界面
