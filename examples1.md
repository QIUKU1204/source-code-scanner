###### GeoBeansƽ̨���ο�����ܹ�����Ҫ�ӿ�˵���ĵ�
### **1.1.1 Բ���غ���̬GLoadObjectCircleShape**
#### **1.1.1.1  GLoadObjectCircleShape**
| �ӿ���ʽ: GLoadObjectCircleShape(QString strID, QString strName, QString strSymbolName) |
| ��������: Բ���غ���̬�����캯�� |
| ����: strID: QString����,�������;Բ���غ���̬����ID��Ϣ<br>	 strName: QString����,�������;Բ���غ���̬����������Ϣ<br>	 strSymbolName: QString����,�������;Բ���غ���̬������ʽ������Ϣ |
| ����ֵ: �� |
| ʾ������: �� |
#### **1.1.1.2  GLoadObjectCircleShape~**
| �ӿ���ʽ: ~GLoadObjectCircleShape() |
| ��������: Բ���غ���̬������������ |
| ����: �� |
| ����ֵ: �� |
| ʾ������: �� |
#### **1.1.1.3  GetCenter**
| �ӿ���ʽ:  GetCenter() |
| ��������: ��ȡԲ���غ���̬��Բ��λ�� |
| ����: �� |
| ����ֵ: QVector3D����: ��ǰԲ�غ���̬��Բ��λ�� |
| ʾ������: �� |
#### **1.1.1.4  SetCenter**
| �ӿ���ʽ: long SetCenter(QVector3D point) |
| ��������: ����Բ���غ���̬��Բ�ľ�γ��λ�� |
| ����: point: QVector3D����,�������; Բ�ľ�γ��λ����Ϣ |
| ����ֵ: long����: ������Բ����Ϣ�ɹ�����0;������Բ����Ϣʧ�ܷ���1 |
| ʾ������: �� |
#### **1.1.1.5  GetRadius**
| �ӿ���ʽ: double GetRadius() |
| ��������: ��ȡԲ���غ���̬�İ뾶 |
| ����: �� |
| ����ֵ: double����: ��ǰԲ�غ���̬�İ뾶����λΪ�� |
| ʾ������: �� |
#### **1.1.1.6  SetRadius**
| �ӿ���ʽ: long SetRadius(double dRadiusInMeter) |
| ��������: ����Բ���غ���̬�İ뾶 |
| ����: dRadiusInMeter: double����,�������; �뾶����λΪ�� |
| ����ֵ: long����: �����ð뾶��Ϣ�ɹ�����0;�����ð뾶��Ϣʧ�ܷ���1 |
| ʾ������: �� |
#### **1.1.1.7  SetParameters**
| �ӿ���ʽ: long SetParameters(QVector3D center, double dRadiusInMeter) |
| ��������: ����Բ���غ���̬�Ĳ�����Ϣ |
| ����: point: QVector3D����,�������; Բ�ľ�γ��λ����Ϣ<br>	<br>	 dRadiusInMeter: double����,�������; �뾶����λΪ�� |
| ����ֵ: long����: �����ò�����Ϣ�ɹ�����0��ͬʱ���������仯ʱ������shapeChanged�ź�;<br>	           �����ò�����Ϣʧ�ܷ���1 |
| ʾ������: �� |
###### ������ң������Ϣ�������޹�˾
