   M S B A M L   `   `   `  ����  compiladorRiqual*    $DocumentUploader.DebugTerminalWindow   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$  Terminal de Debug��5   	   $	��500��5   	   $	��800��6    � WindowStartupLocation$ CenterOwner=�5   	   ���#1E1E1E���#FFFFFF.��5
      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5      $
��Auto �6   .+�5   
   �� $��0��6   ���#2D2D30 �� Padding$ 15,10q�63   ._�5      �  � ColumnDefinitions 5      �� 5      $��* �6'   �� 5      $
��Auto �6'   .+�5      ��  �� Orientation$ 
Horizontal=�6   .+�5      �� $
��🔍��6    $��16(�6*   $��0,0,10,0q�68   �� 5      $/��)Terminal de Debug - Extração de Autores��6    $��14(�5       $��SemiBold"�6.   5      �� 5      -   5         ��Name$ btnClear��5   )   $��1��6   $��Limpar��5       ���#007ACC
�������$��0q�5"      $
��10,5q�6-    � Cursor$
 Hand��5#      $��12(�6'   5$      5%      �� 5(   
   $��1��6   $
��Auto=�5)      $
��Auto=�5*      ���#1E1E1E$��10q�5,      .��5-      �� -   5         ��Name$ 	txtOutput��5-      $!��Consolas, Monaco, monospace)�5.      $��12(�5/      ���#FFFFFF
������ 	 �� TextWrapping$
	 Wrap=�52      
 �� 
IsReadOnly$

 True��53      $��0q�54      " �� VerticalScrollBarVisibility$
 Auto=�55      $ �� HorizontalScrollBarVisibility$
 Auto=�56      $��Terminal a iniciar...
��57      58      �� 5;   
   $��2��6   ���#2D2D30$
 15,8q�63   ._�5<      �  5=      �� 5>      $��* �6'   �� 5?      $
��Auto �6'   .+�5B      �� -   5         ��Name$ 	txtStatus��5B      $��Pronto��5C      $��12(�5D      ���#CCCCCC5E      �� 5G      -   5        $ txtTimestamp��5H      $��1��5G      $��12(�5I      ���#CCCCCC5J      5K      5L      5M      5N      