   M S B A M L   `   `   `  ����J  FWpfApplication2, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication2WpfApplication2  5   	       WpfApplication2.Window3   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:WpfApplication2     5   	     � Title$  Create New Wallet��5   	   $	��650��6#   $	��750��60   .��5	      � .+�5
   
   �� -   5         ��Name$" LabelCreateWalletMainDisplay��5
      $O��I1. Select Location of Simplewallet.exe
2. Enter Wallet Name and Password��66   $��Center=�6�   $��60��6�   $��169,62,169,495q�6�   $��Center=�6�   $	��404��6�   
���   �$��1q�6  $��Center=�6"  $��18.667(�6F  $��Elephant)�6X  �� 5   
   -   5         ��Name$ TextBoxCreateWalletName��5      $
��Left=�63   $��60��6N   $��169,183,0,0q�6Z    �� TextWrapping$
 Wrap=�6o   $��TextBox��6�   $	��Top=�6�   $	��404��6�   �� 5   
   -   5        $ label1��5      $��Wallet Name��6    $
��Left=�66   $��169,157,0,0q�6Q   $	��Top=�6f   $��79��6~   
���   �$��1q�6�   $��27��6�   �� 5   
   -   5        $ LabelCreateWalletPassword��5      $��Password��63   $
��Left=�6F   $��169,273,0,0q�6a   $	��Top=�6v   $��79��6�   $��27��6�   
���   �$��1q�6�   �� 5   
   -   5         ��Name$ LabelUserWelcome��5      $��	20,20,0,0q�6.   $
��Wrap=�6A    J� RenderTransformOrigin$ 0.512,0.845&�6U   $��14.667(�6y   $��Modern No. 20)�6�   $
��Left=�6�   $
��Bold"�6�   ���	#FF574040$��MxNero Wallet V.1.0.1��6�   $	��149��6  $	��Top=�6  $��23��6+  �� 5   
   -   5         ��Name$ ButtonFolderSelection��5      $��Select Location��60   $
��Left=�6J   $��
597,20,0,0q�6e   $	��Top=�6y   $	��120��6�   $��40��6�   �� 5   
   -   5        $ ButtonCreateWallet��5      $��Create New Wallet��6-   $��Center=�6I   $��246,398,246,171q�6f   $��Center=�6   $	��250��6�   $��50��6�   �� 5   
   -   5         ��Name$ image��5      $i��cC:\Users\Eddy\Documents\Visual Studio 2015\Projects\WpfApplication2\WpfApplication2\Monero-Logo.jpg��6   $��Center=�6�   $��75��6�   $��224,505,226,39q�6�   $��Center=�6�   $	��292��6�   G� 5   
   -	   5         G�Name$ PasswordBox��5      $
��Left=�6+   $��169,300,0,0q�6F   $	��Top=�6[   $	��404��6s   $��60��6   �� 5   
   -
   5        	 ��Name$	 CheckBoxShowSimplewallet��5      $��Show Simplewallet?��65   $
��Left=�6R   $��	20,43,0,0q�6m   $	��Top=�6�   $��20��6�   $	��134��6�   
 d� IsThreeState
 .  6�   5      5      