   M S B A M L   `   `   `  ���� C  ?tempproj, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null#clr-namespace:tempprojtempproj      tempproj.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     (localclr-namespace:tempproj       � Title$)  #Converting Process Automation -hufs��$	��650��$
��1000��.���  � 	Resources  �� 
TargetType �� Property �� Value��%y   '��       �� " ��.���  � ��$ 10,10,10,10q���  � ��
 �������  � ��$ 0q���  � ��$ 25�� � ColumnDefinitions �� $��4* ��� $	��43* � � RowDefinitions �� $
��0.1* ��� $��2* �.+�� $��0��$��2��.+�c� .@��� -    ��Name$ btn_MappingTable��$#��매핑 테이블 수정하기��� $��1��$��2��$��
0,0,-0.4,0q� �� $��1* ��� $��2* �.+�� $��0��.+��� -    ��Name$ scrollv��$
��Auto=�.���� .+��� -   $ btnOpenTemplateFile��$#��업로드 자료 불러오기��$��Center=�$��2,2,2,0q�$	��200���� -   $ btnOpenFile��$ ��급여 자료 불러오기��$��Center=�$��2,2,2,0q�$	��200���� -   $ btnStartExcelWork��$��작업 시작하기��$��Center=�$��2,2,2,0q�$	��200���� -   $ ClearListBox��$ ��작업 리스트 비우기��$��Center=�$��2,2,2,0q�$	��200���� -   	 ��Name$	 AnnouncementTextBlock��$
��Wrap=�$��10,10,10,10q�$
��Auto=�
��������� -   $	 DebugConsoleBlock��$
��Wrap=�$��10,10,10,10q�$
��Auto=�
�������� $��1��.+��� $
��Auto=�.���� $��10, 10, 20, 10q� ��%3   '��       �� " ��.���  � ��$ 0,10,0,0q�.+��� -	   
 ��Name$
 ExcelTemplateViewLabel��$<��6TheZone 업로드 자료 (급여수당일괄업로드)���� -
    ��Name$ ExcelTemplateView��$��Stretch=�$��Visible=��� -   $
 ExcelListViewLabel��$*��$급여자료 작업 대기 리스트���� -   $ ExcelListView��$��Stretch=�$��Visible=����� .R�� .+��� -    �� 
IsEditable .  �� 
IsReadOnly .  �� Text$  회사를 선택하세요.��
 � Tag �� *
PathInfo+���� *	cbItems+�� ���� *
PathInfo+�� -   $
 ExcelWorkEndLabel��$��작업 완료 리스트���� -   $ ExcelWorkEndView��$��Stretch=�$��Visible=��� -   $
 ExcelWorkFailLabel��$��작업 실패 리스트���� -   $ ExcelWorkFailView��$��Stretch=�$��Visible=�