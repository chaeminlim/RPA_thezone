   M S B A M L   `   `   `  ���� C  ?tempproj, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null#clr-namespace:tempprojtempproj      tempproj.MappingTable   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     (localclr-namespace:tempproj       � Title$  MappingTable��$	��550��$	��900�� � 	Resources    MyWaterMarkStyle �� 
TargetType �� Property �� Value �� 
TargetType ��Name ��Name
 � Tag	 �� Triggers
 ]� 
Conditions �� Property �� Value �� 
TargetName P� Property P� Value��%�  &	        �� " ��.���  � �� �� " ��.R�� .+��� 
����������	#FF7D8683$��1q��� $ PART_ContentHost��$��5,0,0,0q�$��Center=��� $ WaterMarkLabel��$��5,0,0,0q�#��z $��Center=�$y�	Collapsed=�
�������$��Arial)�	 ]� 
 ��  � ��$  ��./��� $ WaterMarkLabel�� � y�$ Visible=�P�  � {� .  . ���  � ��
 �iii�.���  � RowDefinitions �� $��1* ��� $	��10* �.+�� $��0��.+��� .+��� -    ��Name$ ClientTypeComboBox������ *	cbItems+ �� SelectedItem �� *SelectedcbItem+�� $0��*위 박스에서 회사를 선택하세요��$��Right=�� $��1�� � ColumnDefinitions �� $��1* ��� $��1* �.+�� $��0�� �� $	��12* ��� $��1* �.+�� $��0��.+�R� -    R�Name$ JsonTreeView��� $��1��.+��� -    ��Name$ BtnDeleteJson��$��
5, 5, 5, 5q�$��삭제하기��{�.  � $��1��.+��� $��Visible=�.@��� -    ��Name$ EditCompanyTabItem��$��회사추가��{�. .���� -    ��Name$" EditCompanyTabItemStackPanel��.+��� $#��회사명을 입력하세요.���� -    ��Name$ CompanyNameTextBox��
�������$��4q��� -   $ EditCompanyYesButton��$��추가��$��5,5,5,5q��� -   $ EditCompanyNoButton��$��취소��$��5,5,5,5q��� -	   $ EditSheetTabItem��$��시트추가��{�.  .���� -
   $  EditSheetTabItemStackPanel��.+��� $'��!시트 이름을 입력하세요.���� -   $ SheetNameTextBox��
�������$��4q��� -    ��Name$ SheetFinalInfoTextBlock��
��������� -   $ EditSheetYesButton��$��추가��$��5,5,5,5q��� -   $ EditSheetNoButton��$��취소��$��5,5,5,5q��� -   $ EditMappingTabItem��$��배치정보추가��{�.  $��Visible=�.���� $��
1,64,-1,-2q�.���� -   $" EditMappingTabItemStackPanel�� �� CanVerticallyScroll . $��Visible=�.+��� $=��7구분이 필요하면 체크박스를 선택하세요.���� -    ��Name$ divisionCheckBox���� $'��!시트 이름를 입력하세요.���� -   $ sheetNameTextBox��
�������$��4q��� $$��셀 위치를 입력하세요.���� -   $ cellPointTextBox��
�������$��4q��� $$��셀 이름을 입력하세요.���� -   $ cellNameTextBox��
�������$��4q��� $'��!더존 이름을 입력하세요.���� -   $ theZoneTrueTextBox��
�������$��4q��� -   $ btnTheZoneTrueListBox��$ ��더존 대상 추가하기��
��������� -   $  btnClearTheZoneTrueListBox��$��모두 삭제��
��������� -    ��Name$ theZoneTrueListBox��
�������$��4q�$��5,5,5,5q��� $ ��구분을 입력하세요.���� -   $ divisionTextBox��
��������� $��값을 입력하세요.���� -   $ valueCheckTextBox��
��������� -   $ btnvalueCheckListBox��$��값 대상 추가하기��
��������� -   $ btnClearvalueCheckListBox��$��모두 삭제��
��������� -   $ valueCheckListBox��
�������$��4q�$��5,5,5,5q��� $��False���� -   $ theZoneFalseTextBox��
��������� -   $ btnTheZoneFalseListBox��$��False 대상 추가하기��
��������� -    $! btnClearTheZoneFalseListBox��$��모두 삭제��
��������� -!   $ theZoneFalseListBox��
�������$��4q�$��5,5,5,5q��� -"   $ MappingFinalInfoTextBlock��
��������� -#   $ EditMappingYesButton��$��추가��$��5,5,5,5q��� -$   $ EditMappingNoButton��$��취소��$��5,5,5,5q�