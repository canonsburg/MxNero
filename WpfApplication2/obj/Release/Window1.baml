   M S B A M L   `   `   `  ���� J  FWpfApplication2, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:WpfApplication2WpfApplication2      WpfApplication2.Window1   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     /localclr-namespace:WpfApplication2     -     � Title$  Monero Wallet Manager��$	��700��$	��800��.���  J� RenderTransformOrigin$ 0.519,0.465&� � ColumnDefinitions ��  � RowDefinitions �� �� $
��639* �.+��� -    ��Name$ LabelUserWelcome��$��	20,20,0,0q�$
��Wrap=�$ 0.512,0.845&�$��14.667(�$��Modern No. 20)�$
��Left=�$
��Bold"����	#FF574040$��MxNero Wallet V.1.0.1��$	��149��$	��Top=�$��1��$��23��F� �����P@����	���	#FFB0B0C5$��Right=�$��0��$��0,-18.497,-51.082,0q� �� Stretch$
 Fill=�
���   � � UseLayoutRounding .  $	��Top=�$��0��F� ���Ѐ��� ��	���	#FFB0B0C5$
��Left=�$��0��$��-102,64,0,0q�$
 Fill=�
���   � .  $	��Top=�$��0��$��1��G� -    G�Name$ TextEnterPassword��$��203,275,208,363q�$��30��$��1��$��0q�
���   ��� -    ��Name$ TextWalletPath��$	��100��$��143,149,143,0q�	 �� TextWrapping$
	 Wrap=�$�� ��$	��Top=�$��1��$��2q�
 �� IsReadOnlyCaretVisible
 .  �� 
IsReadOnly . �� -   $ 	textBlock��$��48��$��180,64,180,0q�$��1��$
��Wrap=�$	��Top=�$��18.667(��� $,�&                   Welcome Monero User���� �� $4�.       Please Select Your Wallet Path Location���� -    ��Name$ ButtonSetWalletPath��$��Select Wallet Path��$
��Left=�$��33��$��
648,22,0,0q�$��1��$	��Top=�$��98���� -    ��Name$ image��$i��cC:\Users\Eddy\Documents\Visual Studio 2015\Projects\WpfApplication2\WpfApplication2\Monero-Logo.jpg��$
��Left=�$	��122��$��143,413,0,0q�$��1��$	��Top=�$	��496���� ���	#FFABADB3$��2q�$
��Left=�$��38��$��203,267,0,0q�$��1��$	��Top=�$	��338���� -   $ ButtonEnterWallet��$��Enter Wallet��$
��Left=�$��40��$	��Top=�$	��130��$��309,331,0,0q�$��1���� -	    ��Name$ CheckBoxShowSimplewallet��$��Show Simplewallet.exe?��$
��Left=�$��	20,41,0,0q�$��1��$	��Top=�$��25��$	��147�� d� IsThreeState .  �� -
   $ ButtonNewWallet��$��
New Wallet��$
��Left=�$��
648,64,0,0q�$��1��$	��Top=�$��98��$��30��