   M S B A M L   `   `   `  ����B  >WpfApp1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:WpfApp1WpfApp1  5   	       WpfApp1.MainWindow   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   'localclr-namespace:WpfApp1     5   	     � Title$  
MainWindow��5   	   $	��418��6   $	��648��6)   #��� ��65   .��5	      � $��0,0,0,-6q�6    � ColumnDefinitions 5
   
   �� 5      $	��637 �6   .+�5   
   �� -   5         ��Name$ text1��5      $
��Left=�6#   $��	71,93,0,0q�6>   $
��Wrap=�6Q   $��Input Excel file��6e   $	��Top=�6}   $	��109��6�   �� 5   
   -   5         ��Name$ 	inputText��5      $��34��6%   $��
71,114,0,0q�61    �� TextWrapping$
 Wrap=�6E   $��Input file path here ...��6Y   $	��Top=�6y   $
��Left=�6�   $	��419��6�   $��Italic$�6�   �� 5   
   -   5        $ text2��5      $
��Left=�6#   $��
71,191,0,0q�6>   $
��Wrap=�6R   $��Output path��6f   $	��Top=�6y   $	��109��6�   �� 5   
   -   5        $ 
outputText��5      $
��Left=�6&   $��35��6A   $��
71,208,0,0q�6M   $
 Wrap=�6a   $!��Output folder path here ...��6u   $	��Top=�6�   $	��419��6�   $��Italic$�6�   �� 5   
   -   5         ��Name$ slider��5      $��71,291,110,0q�6!   $	��Top=�67    � Maximum$	 200Q�6O    � Value$ 50Q�6|    � SmallChange$ 1Q�6�   	 � LargeChange$	 10Q�6�   $��29��6�   
 �� IsSnapToTickEnabled$

 True��6�   z�5      V� 5      .~�5      �� �� 5      �� 5      T� 5      5      5      �� 5   
   -   5        $ 	sliderVal��5      $
��Left=�6'   $��548,291,0,0q�6B   $
��Wrap=�6W   $��50��6k   $	��Top=�6u   $��17��6�   $��34��6�   �� 5   
   -   5         ��Name$ inButton��5      $��Select from ...��6#   $
��Left=�6=   $��510,114,0,0q�6X   $	��Top=�6m   $��34��6�   $��77��6�   6�   �� 5   
   -   5        $ exButton��5      $��Export to ...��6#   $��510,208,50,0q�6;   $	��Top=�6Q   $��34��6i   6u   �� 5   
   -	   5        $ 	runButton��5      $	��RUN��6$   $��456,337,37,0q�62   $	��Top=�6H   $��32��6`   6l   �� 5   
   -
   5        $ debug��5      $
��Left=�6#   $��	76,28,0,0q�6>   $
��Wrap=�6Q   $	��Top=�6e   $��44��6}   $	��414��6�   5!      5"      