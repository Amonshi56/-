
# SQL запрос
request = '''&amp;?!имя КАК ?!имя'''

# Шаблон набора данных
dataset = '''		<field xsi:type="DataSetFieldField">
			<dataPath>?!имя</dataPath>
			<field>?!имя</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>?!вопрос</v8:content>
				</v8:item>
			</title>
			<appearance>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>МинимальнаяШирина</dcscor:parameter>
					<dcscor:value xsi:type="xs:decimal">40</dcscor:value>
				</dcscor:item>
			</appearance>
			<inputParameters/>
		</field>'''

# Шаблон вычисляемых выражений
calculated = '''	<calculatedField>
		<dataPath>?!путь</dataPath>
		<expression>?!код</expression>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>?!путь</v8:content>
			</v8:item>
		</title>
		<appearance/>
		<inputParameters/>
	</calculatedField>'''

# Шаблон параметры
parameter = '''	<parameter>
		<name>?!имя</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>?!вопрос</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:boolean</v8:Type>
		</valueType>
		<value xsi:nil="true"/>
		<useRestriction>true</useRestriction>
	</parameter>'''

# Шаблон выборки(нужен для настроек)
selected = '''				<dcsset:item xsi:type="dcsset:SelectedItemField">
					<dcsset:field>?!имя</dcsset:field>
				</dcsset:item>'''

# Шаблон настроек
settings = '''			<dcsset:item xsi:type="dcsset:StructureItemGroup">
				<dcsset:groupItems>
					<dcsset:item xsi:type="dcsset:GroupItemField">
						<dcsset:field>?!имя</dcsset:field>
						<dcsset:groupType>Items</dcsset:groupType>
						<dcsset:periodAdditionType>None</dcsset:periodAdditionType>
						<dcsset:periodAdditionBegin xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionBegin>
						<dcsset:periodAdditionEnd xsi:type="xs:dateTime">0001-01-01T00:00:00</dcsset:periodAdditionEnd>
					</dcsset:item>
				</dcsset:groupItems>
				<dcsset:order>
					<dcsset:item xsi:type="dcsset:OrderItemAuto"/>
				</dcsset:order>
				<dcsset:selection>
					<dcsset:item xsi:type="dcsset:SelectedItemAuto"/>
				</dcsset:selection>
				<dcsset:conditionalAppearance>
					<dcsset:item>
						<dcsset:selection>
							<dcsset:item>
								<dcsset:field>?!имя</dcsset:field>
							</dcsset:item>
						</dcsset:selection>
						<dcsset:filter/>
						<dcsset:appearance>
							<dcscor:item xsi:type="dcsset:SettingsParameterValue">
								<dcscor:parameter>ЦветФона</dcscor:parameter>
								<dcscor:value xsi:type="v8ui:Color">auto</dcscor:value>
							</dcscor:item>
						</dcsset:appearance>
					</dcsset:item>
					<dcsset:item>
						<dcsset:selection>
							<dcsset:item>
								<dcsset:field>?!имя</dcsset:field>
							</dcsset:item>
						</dcsset:selection>
						<dcsset:filter/>
						<dcsset:appearance>
							<dcscor:item xsi:type="dcsset:SettingsParameterValue">
								<dcscor:parameter>ЦветФона</dcscor:parameter>
								<dcscor:value xsi:type="v8ui:Color">style:SpecialTextColor</dcscor:value>
							</dcscor:item>
						</dcsset:appearance>
					</dcsset:item>
				</dcsset:conditionalAppearance>
				<dcsset:outputParameters/>
			</dcsset:item>''' 

# Шаблон корня
root = '''<?xml version="1.0" encoding="UTF-8"?>
<DataCompositionSchema xmlns="http://v8.1c.ru/8.1/data-composition-system/schema" xmlns:dcscom="http://v8.1c.ru/8.1/data-composition-system/common" xmlns:dcscor="http://v8.1c.ru/8.1/data-composition-system/core" xmlns:dcsset="http://v8.1c.ru/8.1/data-composition-system/settings" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<dataSource>
		<name>ИсточникДанных1</name>
		<dataSourceType>Local</dataSourceType>
	</dataSource>
	<dataSet xsi:type="DataSetQuery">
		<name>НаборДанных1</name>
<!--ВставкаШаблонаНабораДанных-->
		<dataSource>ИсточникДанных1</dataSource>
		<query>Выбрать
<!--ВставкаЗапроса--></query>
	</dataSet>
<!--ВставкаШаблонаВычесляемыхВыражений-->
	<parameter>
		<name>ВремяЗавершения</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Время завершения</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:dateTime</v8:Type>
			<v8:DateQualifiers>
				<v8:DateFractions>DateTime</v8:DateFractions>
			</v8:DateQualifiers>
		</valueType>
		<value xsi:type="xs:dateTime">0001-01-01T00:00:00</value>
		<useRestriction>false</useRestriction>
	</parameter>
	<parameter>
		<name>Заказ</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Заказ</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type xmlns:d4p1="http://v8.1c.ru/8.1/data/enterprise/current-config">d4p1:DocumentRef.ЗаказНаПроизводство</v8:Type>
		</valueType>
		<value xsi:type="dcscor:DesignTimeValue">Документ.ЗаказНаПроизводство.ПустаяСсылка</value>
		<useRestriction>false</useRestriction>
	</parameter>
	<parameter>
		<name>НомерСерии</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Номер серии</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:string</v8:Type>
			<v8:StringQualifiers>
				<v8:Length>0</v8:Length>
				<v8:AllowedLength>Variable</v8:AllowedLength>
			</v8:StringQualifiers>
		</valueType>
		<value xsi:type="xs:string"/>
		<useRestriction>false</useRestriction>
	</parameter>
	<parameter>
		<name>Ответственный</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Ответственный</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type xmlns:d4p1="http://v8.1c.ru/8.1/data/enterprise/current-config">d4p1:CatalogRef.Пользователи</v8:Type>
		</valueType>
		<value xsi:type="dcscor:DesignTimeValue">Справочник.Пользователи.ПустаяСсылка</value>
		<useRestriction>false</useRestriction>
	</parameter>
<!--ВставкаШаблонаПараметров-->
	<parameter>
		<name>Допуск</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Допуск</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:decimal</v8:Type>
			<v8:NumberQualifiers>
				<v8:Digits>3</v8:Digits>
				<v8:FractionDigits>3</v8:FractionDigits>
				<v8:AllowedSign>Any</v8:AllowedSign>
			</v8:NumberQualifiers>
		</valueType>
		<value xsi:type="xs:decimal">0.02</value>
		<useRestriction>true</useRestriction>
	</parameter>
	<parameter>
		<name>ТолщинаЗаготовки</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>Толщина заготовки</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:decimal</v8:Type>
			<v8:NumberQualifiers>
				<v8:Digits>5</v8:Digits>
				<v8:FractionDigits>3</v8:FractionDigits>
				<v8:AllowedSign>Any</v8:AllowedSign>
			</v8:NumberQualifiers>
		</valueType>
		<value xsi:type="xs:decimal">0</value>
		<useRestriction>false</useRestriction>
	</parameter>
	<settingsVariant>
		<dcsset:name>Основной</dcsset:name>
		<dcsset:presentation xsi:type="xs:string">Основной</dcsset:presentation>
		<dcsset:settings xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows">
			<dcsset:selection>
<!--ВставкаШаблонаВыборки-->
			</dcsset:selection>
			<dcsset:outputParameters>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>МакетОформления</dcscor:parameter>
					<dcscor:value xsi:type="xs:string">ОформлениеОтчетовЧерноБелый</dcscor:value>
				</dcscor:item>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>РасположениеРесурсов</dcscor:parameter>
					<dcscor:value xsi:type="dcsset:DataCompositionResourcesPlacement">Vertically</dcscor:value>
				</dcscor:item>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:use>false</dcscor:use>
					<dcscor:parameter>Заголовок</dcscor:parameter>
					<dcscor:value xsi:type="v8:LocalStringType"/>
				</dcscor:item>
			</dcsset:outputParameters>
<!--ВставкаШаблонаНастроек-->
		</dcsset:settings>
	</settingsVariant>
</DataCompositionSchema>'''                                       
