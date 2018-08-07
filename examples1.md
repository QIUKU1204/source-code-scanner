###### GeoBeans平台二次开发库架构及主要接口说明文档
### **1.1.1 圆形载荷形态GLoadObjectCircleShape**
#### **1.1.1.1  GLoadObjectCircleShape**
| 接口形式: GLoadObjectCircleShape(QString strID, QString strName, QString strSymbolName) |
| 功能描述: 圆形载荷形态对象构造函数 |
| 参数: strID: QString类型,输入参数;圆形载荷形态对象ID信息<br>	 strName: QString类型,输入参数;圆形载荷形态对象名称信息<br>	 strSymbolName: QString类型,输入参数;圆形载荷形态对象样式名称信息 |
| 返回值: 无 |
| 示例代码: 无 |
#### **1.1.1.2  GLoadObjectCircleShape~**
| 接口形式: ~GLoadObjectCircleShape() |
| 功能描述: 圆形载荷形态对象析构函数 |
| 参数: 无 |
| 返回值: 无 |
| 示例代码: 无 |
#### **1.1.1.3  GetCenter**
| 接口形式:  GetCenter() |
| 功能描述: 获取圆形载荷形态的圆心位置 |
| 参数: 无 |
| 返回值: QVector3D类型: 当前圆载荷形态的圆心位置 |
| 示例代码: 无 |
#### **1.1.1.4  SetCenter**
| 接口形式: long SetCenter(QVector3D point) |
| 功能描述: 设置圆形载荷形态的圆心经纬度位置 |
| 参数: point: QVector3D类型,输入参数; 圆心经纬度位置信息 |
| 返回值: long类型: 若设置圆心信息成功返回0;若设置圆心信息失败返回1 |
| 示例代码: 无 |
#### **1.1.1.5  GetRadius**
| 接口形式: double GetRadius() |
| 功能描述: 获取圆形载荷形态的半径 |
| 参数: 无 |
| 返回值: double类型: 当前圆载荷形态的半径，单位为米 |
| 示例代码: 无 |
#### **1.1.1.6  SetRadius**
| 接口形式: long SetRadius(double dRadiusInMeter) |
| 功能描述: 设置圆形载荷形态的半径 |
| 参数: dRadiusInMeter: double类型,输入参数; 半径，单位为米 |
| 返回值: long类型: 若设置半径信息成功返回0;若设置半径信息失败返回1 |
| 示例代码: 无 |
#### **1.1.1.7  SetParameters**
| 接口形式: long SetParameters(QVector3D center, double dRadiusInMeter) |
| 功能描述: 设置圆形载荷形态的参数信息 |
| 参数: point: QVector3D类型,输入参数; 圆心经纬度位置信息<br>	<br>	 dRadiusInMeter: double类型,输入参数; 半径，单位为米 |
| 返回值: long类型: 若设置参数信息成功返回0，同时参数发生变化时将发出shapeChanged信号;<br>	           若设置参数信息失败返回1 |
| 示例代码: 无 |
###### 北京中遥地网信息技术有限公司
