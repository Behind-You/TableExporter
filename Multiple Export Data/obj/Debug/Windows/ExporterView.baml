   M S B A M L   `   `   `  ����O  KMultiple Export Data, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullK*clr-namespace:Multiple_Export_Data.WindowsMultiple_Export_Data.Windows  5      /    )Multiple_Export_Data.Windows.ExporterView   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      g cMahApps.Metro.IconPacks.Material, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0c0d510f9915137ac _MahApps.Metro.IconPacks.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0c0d510f9915137a:Icon-http://metro.mahapps.com/winfx/xaml/iconpacks   5      Fmc;http://schemas.openxmlformats.org/markup-compatibility/2006   5      8local*clr-namespace:Multiple_Export_Data.Windows   5        N� Title$  ExporterView��5      .��5      � $��
30,0,20,20q�6    � RowDefinitions 5	   
   �� 5
      $��50 �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $��* �6   �� 5      $
��auto �6   .+�5   
   �� $��Table Exporter��6   $��0��6*   $��32(�67   $��SemiBold"�6E   $��0,0,0,0q�6[   $
��Left=�6l   ���#121518$��Bottom=�6�   �� 5   
   -   5         ��Name$ ExportButton��5         	AddButton#��[  6'   $��0��6J   .��5      ��  �� Orientation$ 
Horizontal=�6   .+�5      .  (MahApps.Metro.IconPacks.PackIconMaterial     Kind$ Database=�6(   $��Center=�68   $��11��6S   $��11��6^   $��0,1,8,0q�6j   �� 5      $��Export��6   5      5      �� 5   
   $��1��6   $��2��6   ���#dae2ea$��0,0,0,0q�6A   �� 5   
   -   5         ��Name$ TextBlock_DataGridTitle��5      $��Export Settings��65   $��3��6L   $��18(�6Y   $��SemiBold"�6g   ���#121518$
��Left=�6�   $��0,0,0,2q�6�   �� 5"   
   $��1��6   $��4��6   ���#dae2ea$��0,0,0,0q�6A   � 5#   
   $��3��6   $��0,0,0,0q�6   $��Right=�6-    � ColumnDefinitions 5$      �� 5%      $
��auto �6#   �� 5&      $
��auto �6#   �� 5'      $
��auto �6#   �� 5(      $
��auto �6#   .+�5*      �� $��0��6   $ 
Horizontal=�6)   $��10,0,0,0q�6B   .+�5+      �� $��Server Type��6   $��14(�6/   $��SemiBold"�6=   ���#121518$
��Left=�6h   �� 5,      -   5         ��Name$ ComboBox_Filter_Server��5,      $��10,0,0,0q�6;   $	��100��6M   5-      �� 5.      $��1��6   $ 
Horizontal=�6)   $��10,0,0,0q�6B   .+�5/      �� $��Legion Type��6   $��14(�6/   $��SemiBold"�6=   ���#121518$
��Left=�6h   �� 50      -   5        $ ComboBox_Filter_Legion��50      $��10,0,0,0q�6;   $	��100��6M   51      �� 52      -   5        $ Btn_Refresh_Filter��52      $��2��61     SystemButton#��[ 6A   $��0q�6g   .��53        $ TableSearch=�6(   4  .System.Windows.Automation.AutomationProperties   HelpText$ Serch With Filter��6;   54      �� 55      -   5        $ Btn_Remove_Filter��55      $��3��60   #��[ 6@   $��0q�6f   $ Reset Filter��6�   .��56        $ Refresh=�6(   57      58      &   System.Windows.Controls.DataGrid  5:   
   -   5        	  Name$	 MainDataGrid��5:   �     DataGridStyle1#���  6     DataGridCellStyle1
   	CellStyle#
 �  6<     DataGridColumnHeaderStyle1   ColumnHeaderStyle# �  6m     DataGridRowStyle1   RowStyle# �  6�   $��5��6�   $��-2,2,2,0q�6�      Columns 5;      4  .System.Windows.Controls.DataGridCheckBoxColumn  5<      -   5        '#DataGridCheckBoxColumn_ExportTarget��,  &System.Windows.Controls.DataGridColumn   CanUserResize .     Width5  /System.Windows.Controls.DataGridLengthConverter$
 auto 5<   l   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6y   �� *
IsSelected+   HeaderTemplate 5=      �� 5>      .R�5?      �� -	   5        5?   '   5@      5B      0  *System.Windows.Controls.DataGridTextColumn  5D         Header$ #��6%      
IsReadOnly .  .  $
 auto 6X    6e   �� *INDEX+  5E      $ Server Types��6%    . $
 auto 6M    6Z   �� *SERVER_TYPE+  5F      $ Legion Types��6%    . $
 auto 6M    6Z   �� *LEGION_TYPE+  5G      $
 Name��6%    . $
 auto 6E    6R   �� *NAME+  5H      $
 Path��6%    . $
 auto 6E    6R   �� *PATH+5K      5L      5M      