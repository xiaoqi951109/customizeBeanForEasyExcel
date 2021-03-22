# customizeBeanForEasyExcel
Easyexcel mainly supports Java annotation to define objects, but there is no way for non spring system or Java Bean scheme. For example, the JSON format. Therefore, I use the exposed converter interface and writerstrategy interface to simply encapsulate the custom cell object, so that all custom types can use this code to use easyexcel

# 场景
对于不方便使用JavaBean 和注解来进行Excel的操作，同时poi基础包的性能又不满足业务需求的情况下，可以考虑使用自定义的通用Cell，三方的数据/对象/JSON API 转化成Cell的方式进行EasyExcel的操作使用。

# 特性
支持跨行跨列的快速设置、所有的样式都可以基于单个Cell。

# 扩展
目前只是简单按需支持了部分样式、格式，但是代码比较简单，所有的样式和格式都可以支持，再对应的Cell类里面添加属性，在Converter里面进行类型设置，然后在Strategy里面进行样式处理就可以了。

# 代码风格
个人习惯使用 属性  B.a(v)、B.a() 的方式搭配builder模式进行数据的插入和获取。有需要或者不习惯的话可以根据自己的情况进行修改。
