   M S B A M L   `   `   `  ���� D  @irRenamer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:irRenamer	irRenamer      irRenamer.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     )localclr-namespace:irRenamer       � Title$  不规律的重命名��$	��560��$	��400�� � 
ResizeMode$ NoResize=� � WindowStartupLocation$ CenterScreen=�.��� .+���  �� Padding$ 10q�._��� .+��� $��1. 选择文件��$
��Bold"�$��0 0 0 0q��� $T��N任意选择一个需要重命名的文件，用于获取需要处理的目录��
�������$��0 5 0 5q�� $��24�� � ColumnDefinitions �� $��* ��� $��60 �.+��� -   $��TextBox_Path��$��0��$��Center=�$��0 0 5 0q��� -   $��
Select_Btn��$��选择文件��$��1��
����ح��� $��2. 获取文件名��$
��Bold"�$��0 10 0 0q��� $K��E获取目录下的所有文件和文件夹（不包含子文件夹）��
�������$��0 5 0 0q��� $J��D如果需要保留之前生成的Excel文件，可以跳过此步骤��
�������$��0 0 0 5q��� -   $��GetNames_Btn��$��执行操作��$��24��
����ح�${�False���� $��3. 修改文件名��$
��Bold"�$��0 10 0 0q��� $V��P在打开的Excel文件中修改重命名后的名称（不填表示不修改）��
�������$��0 5 0 0q��� $2��,修改完成后请保存并关闭Excel文件��
�������$��0 0 0 5q��� -   $��OpenExcel_Btn��$��打开Excel文件��$��24��
����ح�${�False���� $��4. 重命名��$
��Bold"�$��0 10 0 0q��� $5��/按照Excel中的文件名进行重命名操作��
�������$��0 5 0 0q��� $N��H不需要重命名的在新文件名和新扩展名单元格留空即可��
�������$��0 0 0 5q��� -   $��
Rename_Btn��$��	重命名��$��24��
����ح�${�False���� $��0 60 0 0q��� $��	说明：��$
��Bold"�$��0 10 0 0q��� $B��<本工具主要用于没有规律的批量重命名操作。��
�������$��0 5 0 0q��� $N��H请按照上述顺序依次操作，避免产生不可预知的错误。��
�������$��0 5 0 0q��� $]��W本工具也适用于有规律的批量重命名，不过您应该有更好的选择。��
�������$��0 5 0 0q�