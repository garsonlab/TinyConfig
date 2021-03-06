# TinyConfig
Excel快速转换CSV工具，并生成CSharp读取配置文件。

包含功能： 转客户端csv，客户端读取配置的CSharp代码；转服务器csv，后续支持直转数据库

----


### Excel表头格式
| 输出  | OtherDef  |   |   |   |   |   |   |   |   |
| ------------ | ------------ | ------------ | ------------ | ------------ | ------------ | ------------ | ------------ | ------------ | ------------ |
|  服务器 |   | Level  |   | Exp  | Max  | Property  | Type  |  Name | Model  |
| 客户端  | Id  | Level  |  Name | Exp#Min  | Exp#Max  | Property#1  | Property#2  | Monster#1#1#Name  |  Monster#1#2#Name |
|  类型 | _key  | _key  | text  |  int | int  | int  | int  |  text |  text |
| 说明  | id  | 等级  | 名字  |  经验 | 最大经验  | 属性  | 类型  | 名字1  | 名字2  |

#### 字段说明
* 输出：生成的csv\cs文件、类名
* 服务器：标识服务器使用字段，不填不会转换
* 客户端：同服务器
* 类型：支持 text(string), int, byte, long, float, double
* _key指主键，用于查找配置表使用，至少1个，至多3个，类型为int

*生成CSharp文件中，字段分割使用“#”，遇到数字自动识别成数组，其他识别成类。数组下标从1开始*


### Converter转换器使用
``` csharp
public struct Options
{
    public string excelPath;//excel
    public string serverFolder;//server csv
    public string clientFolder;//client csv
    public string csFolder;//csharp
    public string nameSpace;//csharp 命名空间
}
```

### 栗子
使用“测试.xlsx”进行转换，根据下方的table生成两个CSharp文件， [CombatExpDef](./TinyConfig/Test/CombatExpDef.cs), [OtherExpDef](./TinyConfig/Test/OtherExpDef.cs)

每个类包含3个静态模块:
> * Load(string)，加载读取的csv数据
> * Values，所有def的List
> * Find(int ...), 根据上方标识的“_key”获取单个配置项

### TODO
* 支持转义
* 支持枚举
* 支持配置直接索取