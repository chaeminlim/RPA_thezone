   M S B A M L   `   `   `  ���� C  ?tempproj, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null#clr-namespace:tempprojtempproj      tempproj.MappingTable   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     (localclr-namespace:tempproj       � Title$  MappingTable��$	��550��$	��900�� � 	Resources    MyWaterMarkStyle �� 
TargetType �� Property �� Value �� 
TargetType ��Name ��Name
 � Tag	 �� Triggers
 ]� 
Conditions �� Property �� Value �� 
TargetName P� Property P� Value��%�  &	        �� " ��.���  � �� �� " ��.R�� .+��� 
����������	#FF7D8683$��1q��� $ PART_ContentHost��$��5,0,0,0q�$��Center=��� $ WaterMarkLabel��$��5,0,0,0q�#��z $��Center=�$y�	Collapsed=�
�������$��Arial)�	 ]� 
 ��  � ��$  ��./��� $ WaterMarkLabel�� � y�$ Visible=�P�  � {� .  . ���  � ��
 �iii�.���  � RowDefinitions �� $��1* ��� $	��10* �.+�� $��0��.+��� .+��� -    ��Name$ ClientTypeComboBox������ *	cbItems+ �� SelectedItem �� *SelectedcbItem+�� $0��*위 박스에서 회사를 선택하세요��$��Right=�� $��1�� � ColumnDefinitions �� $��1* ��� $��1* �.+��� $��0��$��0,0,-0.4,-0.4q�.+�&   System.Windows.Controls.DataGrid  -     Name$ MappingDataGrid��   Columns 4  .System.Windows.Controls.DataGridCheckBoxColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Included��1  +System.Windows.Controls.DataGridBoundColumn   Binding ��  �� Path$ 
IsIncluded�0  *System.Windows.Controls.DataGridTextColumn  $
 From�� �� $ 	KeyString�  $ To�� �� $ ValueString�  $
 Type�� �� $ 
TypeString�  $ TypeName�� �� $ TypeNameString�  $
 True�� �� $ 
TrueString�  $ False�� �� $ FalseString��� $��1��.+��� -   $ StatusLabel��$�� ���� $
��From���� -    ��Name$ FromTextBox��#��[  $# Center파일의 항목 이름���� $��To���� -   $ 	ToTextBox������ *TheZoneItems+ �� *SelectedtzItem+�� $
��Type���� -   $ TypeTextBox��#��[  $# 확인해야할 항목 이름���� $��TypeName���� -   $ TypeNameTextBox��#��[  $( "일치해야할 값 (, 로 구분)���� $��일치���� -   $ TrueTextBox��#��[  $4 .일치한다면 연결될 TheZone항목 이름���� $��	불일치���� -	   $ FalseTextBox��#��[  $> 8일치하지 않는다면 연결될 TheZone항목 이름���� -
    ��Name$ 
btn_AddRow��$��추가하기���� -   $ btn_deleteRow��$��선택 지우기���� -   $ btn_Save��$��저장하기��