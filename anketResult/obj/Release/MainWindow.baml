   M S B A M L   `   `   `  ���� F  BanketResult, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:anketResultanketResult      anketResult.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     +localclr-namespace:anketResult       � Title$�  �Составление протоколов анкетирование Алтайского промышленно-экономического колледжа��$	��350��$��710.294�� � 
ResizeMode$ NoResize=� � Icon$ aiec_logo.ico��.���  � RowDefinitions �� $	��32* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��30* ��� $	��55* � � ColumnDefinitions �� $
��205* ��� $
��356* ��� $
��141* �.+��� -    ��Name$ 
SelectSpec��$��1��$
��Left=�$��3��$	��Top=�$	��337��$��20��$��10,0,0,0q� �� SelectedIndex$ 0���� -    ��Name$ FileName��$��1��$
��Left=�$��21��$��5�� �� TextWrapping$
 Wrap=�$	��Top=�$	��337��	 �� TextAlignment$	 Justify=�$��10,0,0,0q�
 �� 	UndoLimit$

 2000���� -   $)��#Протокол (Учащиеся)��$��1��$
��Left=�$��7��$	��Top=�$	��140��$��21��$��98,0,0,0q��� -   $ expert��$��1��$
��Left=�$��20��$��10,0,0,0q�$��1��$
 Wrap=�$��ФИО эксперта��$	��Top=�$	��337���� -   $-��'Протокол (Сотрудники)��$��1��$
��Left=�$��98,0,0,0q�$��8��$	��Top=�$	��140��$��20���� -   $��Все протоколы��$��1��$
��Left=�$��98,0,0,0q�$��9��$	��Top=�$	��140��$��20���� $
��Left=�$��99,1,0,0q�$
��Wrap=�$,��&Введите ФИО эксперта��$	��Top=�$��2��$��96�� �� TextAlignment$ Right=�$��32��$��1���� $
��Left=�$��99,2,0,0q�$
��Wrap=�$1��+Выберите специальность��$	��Top=�$��2��$��96��$ Right=�$��3��$��32���� $
��Left=�$��99,8,0,0q�$
��Wrap=�$!��Название файла��$	��Top=�$��96��$ Right=�$��5��$��16���� -   $!��Просмотр анкет��$
��Left=�$��70,0,0,0q�$��8��$	��Top=�$	��126��