   M S B A M L   `   `   `  ����J  FWpfApplication2, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication2WpfApplication2  5   	       WpfApplication2.Window2   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:WpfApplication2     5   	   -   5          � Title$  Monero Wallet Manager��5   	   $	��650��6'   $	��750��64   $��1q�5
   &   
���   �.��5      �  � ColumnDefinitions 5   
   �� 5      .+�5   
   �� ���	#FFE0E0E0$
��Left=�6%   $	��199��6@   $��	48,25,0,0q�6M   
���   �$	��Top=�6o   $	��646��6�   �� 5   
   ���	#FFE0E0E0$
��Left=�6%   $	��239��6@   $��
49,244,0,0q�6M   
���   �$	��Top=�6p   $	��646��6�   �� 5   
   ���	#FFE0E0E0$
��Left=�6%   $��80��6@   $��
48,505,0,0q�6L   
���   �$	��Top=�6o   $	��347��6�   �� 5   
   ���	#FFE0E0E1$
��Left=�6%   $��80��6@   $��411,505,0,0q�6L   
���   �$	��Top=�6p   $	��184��6�   �� 5   
   -   5         ��Name$ image��5      $n��hC:\Users\Eddy\Documents\Visual Studio 2015\Projects\WpfApplication2\WpfApplication2\Monero-Logo-Only.jpg��6   $
��Left=�6�   $��80��6�   $	��Top=�6�   $��81��6�   $��613,505,0,0q�6�   �� 5   
   -   5         ��Name$ LabelWalletBalance��5      $��Wallet Account��6,   $
��Left=�6E   $��25��6`   $��	49,25,0,0q�6l   $	��Top=�6   $	��100��6�   $��1q�6�   �� 5   
   -   5        $ LabelSendMonero��5      $��Send Monero��6)   $
��Left=�6?   $��25��6Z   $��
49,244,0,0q�6f   $	��Top=�6z   $	��100��6�   $��1q�6�   �� 5   
   -   5         ��Name$ TextAddress��5      $
��Left=�6'   $��25��6B   $��
123,55,0,0q�6N    �� TextWrapping$
 Wrap=�6b   $��TextBox��6v   $	��Top=�6�   $	��561��6�    �� 
IsReadOnly . $��9.333(�6�   �� 5   
   -   5        $ LabelAddress��5      $��Address��6&   $
��Left=�68   $��25��6S   $��	67,55,0,0q�6_   $	��Top=�6r   $��56��6�   $��1q�6�   
���   �6�   �� 5   
   -   5        $ LabelBalance��5      $��Balance��6&   $
��Left=�68   $��25��6S   $��
123,94,0,0q�6_   $	��Top=�6s   $��56��6�   
���   �$��1q�6�   �� 5   
   -   5        $ TextBalance��5      $
��Left=�6'   $��25��6B   $��
179,94,0,0q�6N   $
 Wrap=�6b   $��TextBox��6v   $	��Top=�6�   $	��203��6�    J� RenderTransformOrigin$ 
0.504,0.44&�6�    . $��Right=�6�   �� 5   
   -	   5        $ LabelUnlockedBalance��5      $��Unlocked Balance��6.   $
��Left=�6I   $��25��6d   $��
73,129,0,0q�6p   $	��Top=�6�   $	��106��6�   
���   �$��1q�6�   �� 5   
   -
   5        $ TextUnlockedBalance��5      $
��Left=�6/   $��25��6J   $��179,129,0,0q�6V   $
 Wrap=�6k   $��TextBox��6   $	��Top=�6�   $	��203��6�    . $��Right=�6�   �� 5   
   -   5        $ LabelAvailableBalance��5      $��Available Balance��6/   $
��Left=�6K   $��25��6f   $��
73,168,0,0q�6r   $	��Top=�6�   $	��106��6�   
���   �$��1q�6�   �� 5   
   -   5        $ TextAvailableBalance��5      $
��Left=�60   $��25��6K   $��179,168,0,0q�6W   $
 Wrap=�6l   $��TextBox��6�   $	��Top=�6�   $	��203��6�    . $��Right=�6�   �� 5   
   -   5         ��Name$ 
ButtonSave��5      $��Save Wallet��6%   $
��Left=�6;   $��
448,94,0,0q�6V   $	��Top=�6j   $	��236��6�   $��99��6�   $��24(�6�   ���	#FF706F6F$��2q�6�   ��5      �� 5       	 �� EndPoint$	 0.5,1&�6&   
 �� 
StartPoint$
 0.5,0&�67   .��5!      �  � Color$ 	#FFFBDE9F��6#    � Offset$ 0Q�65   � 5"      $ 	#FFFD5D01��6#   $ 0.993Q�65   5#      5%      �� 5&   
   -   5        $% LabeBoarderOfLabelSendToAddress��5&      $�� ��69   $
��Left=�6D   $��37��6_   $��
67,274,0,0q�6k   $	��Top=�6   $��56��6�   $��1q�6�   
���   �6�   �� 5'   
   -   5         ��Name$ LabelSendToAddress��5'      $
��Left=�60   $��37��6K   $��
73,274,0,0q�6W   $
��Wrap=�6k   $	��Top=�6   $��56��6�   �� 6�   $�Send to Address��6�   �� 6�   �� 6�   6�   �� 5(   
   -   5        $ TextSendToAddress��5(      $
��Left=�6-   $��37��6H   $��123,274,0,0q�6T   $
 Wrap=�6i   $��TextBox��6}   $	��Top=�6�   $	��561��6�   $��9.333(�6�   $
��Left=�6�   �� 5)   
   -   5        $ LabelSendAmount��5)      $��Amount to Send��6)   $
��Left=�6B   $��25��6]   $��
67,327,0,0q�6i   $	��Top=�6}   $	��112��6�   
���   �$��1q�6�   �� 5*   
   -   5        $ TextSendAmount��5*      $
��Left=�6*   $��25��6E   $��179,327,0,0q�6Q   $
 Wrap=�6f   $��TextBox��6z   $	��Top=�6�   $	��203��6�   $��Right=�6�   �� 5+   
   -   5         ��Name$ sliderMixin��5+      $
��Left=�6&   $��38��6A   $��
67,366,0,0q�6M   $	��Top=�6a   $	��190��6y    � Minimum$ 2Q�6�    � Maximum$ 20Q�6�    � LargeChange$ 5Q�6�    � SmallChange$ 1Q�6�    �� TickFrequency$ 1Q�6�    �� TickPlacement$
 Both=�6�    �� IsSnapToTickEnabled .  � Value 6�   �� *Text+ �� ElementName$ TextMixinShowOutput�� �� Mode$ TwoWay=�6-  �� 5,   
   -   5        $ LabelMixinLevel��5,      $��Mixin Level��6)   $
��Left=�6?   $��25��6Z   $��308,357,0,0q�6f   $	��Top=�6{   $��74��6�   
���   �$��1q�6�   �� 5-   
   -   5        $ LabelPaymentID��5-      $��Payment ID (Optional)��6(   $
��Left=�6H   $��30��6c   $��
67,404,0,0q�6o   $	��Top=�6�   $	��131��6�   $��1q�6�   
���   �6�   �� 5.   
   -   5        $ TextPaymentID��5.      $
��Left=�6)   $��30��6D   $��
67,434,0,0q�6P   $
 Wrap=�6d   $��TextBox��6x   $	��Top=�6�   $	��617��6�   $��9.333(�6�   �� 5/   
   -   5        $ ButtonSendMoney��5/      $��Send Monero��6*   $
��Left=�6@   $��89��6[   $��448,327,0,0q�6g   $	��Top=�6|   $	��236��6�   $��24(�6�   $��2q�6�   ��50      �� 51      $	 0.5,1&�6&   $
 0.5,0&�67   .��52      � $ 	#FFFBDE9F��6#   $ 0Q�65   � 53      $ 	#FFFD5D01��6#   $ 1Q�65   54      56      �� 57   
   -   5        $ LabelDaemonAddress��57      $��Daemon Address��6,   $
��Left=�6E   $��30��6`   $��
66,515,0,0q�6l   $	��Top=�6�   $	��112��6�   $��1q�6�   
���   �6�   �� 58   
   -   5        $ TextDaemonAddress��58      $
��Left=�6-   $��30��6H   $��178,515,0,0q�6T   $
 Wrap=�6i   $��TextBox��6}   $	��Top=�6�   $	��199��6�    �� IsReadOnlyCaretVisible . ���	#FFE0E0E06�   �� 59   
   -   5        $ LabelDefaultDaemonAddress��59      $#��Default Remote Daemon Address��63   $
��Left=�6[   $��26��6v   $��
66,550,0,0q�6�   $	��Top=�6�   $	��311��6�   $��1q�6�   
���   �6�   �� 5:   
   -   5        $ LabelMvNeroVersion��5:      $��MxNero Version 1.0.0��6,   $
��Left=�6K   $��30��6f   $��425,513,0,0q�6r   $	��Top=�6�   $	��158��6�   $��1q�6�   
���   �6�   �� 5;   
   -   5        $ LabelLightweightWallet��5;      $��Lightweight Monero Wallet��60   $
��Left=�6T   $��31��6o   $��425,548,0,0q�6{   $	��Top=�6�   $	��158��6�   
���   �$��1q�6�   �� 5<   
   -   5        $ LabelXMRSendTo��5<      $	��XMR��6(   $
��Left=�66   $��25��6Q   $��382,327,0,0q�6]   $	��Top=�6r   $��36��6�   �� 5=   
   -   5        $ LabelXMRAvailBal��5=      $	��XMR��6*   $
��Left=�68   $��25��6S   $��382,168,0,0q�6_   $	��Top=�6t   $��36��6�   �� 5>   
   -   5        $ LabelXMRUnlockBal��5>      $	��XMR��6+   $
��Left=�69   $��25��6T   $��382,129,0,0q�6`   $	��Top=�6u   $��36��6�   �� 5?   
   -    5        $ LabelXMRBalance��5?      $	��XMR��6)   $
��Left=�67   $��25��6R   $��
382,94,0,0q�6^   $	��Top=�6r   $��36��6�   �� 5@   
   -!   5        $ TextMixinShowOutput��5@      $
��Left=�6/   $��25��6J   $��279,357,0,0q�6V   $
 Wrap=�6k   $	��Top=�6�   $��29��6�    . ��6   �� *Value+$ sliderMixin��$ TwoWay=�6�   5B      5C      