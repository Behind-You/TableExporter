   M S B A M L   `   `   `  ����O  KMultiple Export Data, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullK*clr-namespace:Multiple_Export_Data.WindowsMultiple_Export_Data.Windows  5      5    /Multiple_Export_Data.Windows.ExportSettingsView   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      g cMahApps.Metro.IconPacks.Material, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0c0d510f9915137ac _MahApps.Metro.IconPacks.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0c0d510f9915137a:Icon-http://metro.mahapps.com/winfx/xaml/iconpacks   5      Fmc;http://schemas.openxmlformats.org/markup-compatibility/2006   5      8local*clr-namespace:Multiple_Export_Data.Windows   5        N� Title$  ExportSettingsView��5      .��5      � $��
30,0,20,20q�6    � RowDefinitions 5	   
   �� 5
      $��50 �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $��* �6   �� 5      $
��auto �6   .+�5   
   �� $��Export Settings��6   $��0��6+   $��32(�68   $��SemiBold"�6F   $��0,0,0,0q�6\   $
��Left=�6m   ���#121518$��Bottom=�6�   �� 5   
   $��1��6   $��2��6   ���#dae2ea$��0,0,0,0q�6A   �� 5   
   -   5         ��Name$ TextBlock_DataGridTitle��5      $��Export Settings��65   $��3��6L   $��18(�6Y   $��SemiBold"�6g   ���#121518$
��Left=�6�   $��0,0,0,2q�6�   �� 5   
   $��1��6   $��4��6   ���#dae2ea$��0,0,0,0q�6A   � 5   
   $��3��6   $��0,0,0,0q�6   $��Right=�6-    � ColumnDefinitions 5      �� 5      $
��auto �6#   �� 5       $
��auto �6#   �� 5!      $
��auto �6#   �� 5"      $
��auto �6#   .+�5$      �� $��0��6    �� Orientation$ 
Horizontal=�6)   $��10,0,0,0q�6B   .+�5%      �� $��Server Type��6   $��14(�6/   $��SemiBold"�6=   ���#121518$
��Left=�6h   �� 5&      -   5         ��Name$ ComboBox_Filter_Server��5&      $��10,0,0,0q�6;   $	��100��6M   5'      �� 5(      $��1��6   $ 
Horizontal=�6)   $��10,0,0,0q�6B   .+�5)      �� $��Legion Type��6   $��14(�6/   $��SemiBold"�6=   ���#121518$
��Left=�6h   �� 5*      -   5        $ ComboBox_Filter_Legion��5*      $��10,0,0,0q�6;   $	��100��6M   5+      �� 5,      -   5         ��Name$ Btn_Refresh_Filter��5,      $��2��61      SystemButton#��[  6A   $��0q�6g   .��5-      .  (MahApps.Metro.IconPacks.PackIconMaterial     Kind$ TableSearch=�6(   5.      �� 5/      -   5        $ Btn_Remove_Filter��5/      $��3��60   #��[  6@   $��0q�6f   .��50        $ Refresh=�6(   51      52      &   System.Windows.Controls.DataGrid  54   
   -   5          Name$ MainDataGrid��54   �     DataGridStyle1#���  6     DataGridCellStyle1	   	CellStyle#	 �  6<     DataGridColumnHeaderStyle1
   ColumnHeaderStyle#
 �  6m     DataGridRowStyle1   RowStyle# �  6�   $��5��6�   $��-2,2,2,0q�6�      Columns 55      0  *System.Windows.Controls.DataGridTextColumn  56      ,  &System.Windows.Controls.DataGridColumn   Header$ #��6%      
IsReadOnly .    CanUserResize .     Width5  /System.Windows.Controls.DataGridLengthConverter$
 auto 6X   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6e   �� *INDEX+  57      $ Server Types��6%    . $
 auto 6M    6Z   �� *SERVER_TYPE+  58      $ Legion Types��6%    . $
 auto 6M    6Z   �� *LEGION_TYPE+  59      $
 Name��6%    . $
 auto 6E    6R   �� *NAME+  5:      $
 Path��6%    . $
 auto 6E    6R   �� *PATH+5<      5=      5>      