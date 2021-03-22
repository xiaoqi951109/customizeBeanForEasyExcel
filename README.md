# customizeBeanForEasyExcel
Easyexcel mainly supports Java annotation to define objects, but there is no way for non spring system or Java Bean scheme. For example, the JSON format. Therefore, I use the exposed converter interface and writerstrategy interface to simply encapsulate the custom cell object, so that all custom types can use this code to use easyexcel

# 场景
对于不方便使用JavaBean 和注解来进行Excel的操作，同时poi基础包的性能又不满足业务需求的情况下，可以考虑使用自定义的通用Cell，三方的数据/对象/JSON API 转化成Cell的方式进行EasyExcel的操作使用。

