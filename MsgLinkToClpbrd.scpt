FasdUAS 1.101.10   ��   ��    k             l    a ����  O     a  	  k    ` 
 
     l   ��������  ��  ��        l   ��  ��    5 / get the currently selected message or messages     �   ^   g e t   t h e   c u r r e n t l y   s e l e c t e d   m e s s a g e   o r   m e s s a g e s      r        n   	    I    	�������� 0 
item_check 
item_Check��  ��     f      o      ���� $0 selectedmessages selectedMessages      l   ��������  ��  ��        r        I   �� ��
�� .corecnte****       ****  o    ���� $0 selectedmessages selectedMessages��    o      ���� 0 msgcount msgCount       l   ��������  ��  ��      ! " ! Z      # $���� # l    %���� % A     & ' & o    ���� 0 msgcount msgCount ' m    ���� ��  ��   $ L    ����  ��  ��   "  ( ) ( l  ! !��������  ��  ��   )  * + * r   ! ' , - , n   ! % . / . 4  " %�� 0
�� 
cobj 0 m   # $����  / o   ! "���� $0 selectedmessages selectedMessages - o      ���� "0 selectedmessage selectedMessage +  1 2 1 l  ( (��������  ��  ��   2  3 4 3 Z   ( F 5 6���� 5 =  ( - 7 8 7 n   ( + 9 : 9 m   ) +��
�� 
pcls : o   ( )���� "0 selectedmessage selectedMessage 8 m   + ,��
�� 
cEvt 6 k   0 B ; ;  < = < I  0 ?�� > ?
�� .sysodlogaskr        TEXT > m   0 1 @ @ � A A � S e l e c t e d   m e s s a g e   i s   a n   E v e n t   /   R e m i n d e r .   P l e a s e   s e l e c t   a n   E - M a i l   m e s s a g e . ? �� B C
�� 
disp B m   2 3����  C �� D E
�� 
btns D m   4 5 F F � G G  O K E �� H I
�� 
dflt H m   6 7 J J � K K  O K I �� L��
�� 
givu L m   8 9���� ��   =  M�� M L   @ B����  ��  ��  ��   4  N O N l  G G��������  ��  ��   O  P Q P r   G T R S R c   G P T U T n   G L V W V 1   H L��
�� 
ID   W o   G H���� "0 selectedmessage selectedMessage U m   L O��
�� 
TEXT S o      ���� 0 msgid msgID Q  X Y X l  U U��������  ��  ��   Y  Z�� Z I  U `�� [��
�� .JonspClpnull���     **** [ b   U \ \ ] \ m   U X ^ ^ � _ _  o u t l o o k : / / ] o   X [���� 0 msgid msgID��  ��   	 m      ` `�                                                                                  OPIM  alis    x  Macintosh HD               ��F�H+   f�@Microsoft Outlook.app                                          �H��$        ����  	                Applications    ��*�      ���     f�@  0Macintosh HD:Applications: Microsoft Outlook.app  ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  ��  ��     a b a l     ��������  ��  ��   b  c d c l     ��������  ��  ��   d  e f e l     �� g h��   g   get selected items    h � i i &   g e t   s e l e c t e d   i t e m s f  j k j i      l m l I      �������� 0 
item_check 
item_Check��  ��   m k     � n n  o p o l     �� q r��   q ) #set myPath to (path to home folder)    r � s s F s e t   m y P a t h   t o   ( p a t h   t o   h o m e   f o l d e r ) p  t�� t O     � u v u k    � w w  x y x Q    � z {�� z l   � | } ~ | k    �    � � � r     � � � 1    ��
�� 
sele � o      ���� 0 selecteditems selectedItems �  � � � r     � � � l    ����� � n     � � � m    ��
�� 
pcls � o    ���� 0 selecteditems selectedItems��  ��   � o      ���� 0 	raw_class 	raw_Class �  � � � Z    V � ����� � =    � � � o    ���� 0 	raw_class 	raw_Class � m    ��
�� 
list � k    R � �  � � � r    ! � � � J    ����   � o      ���� 0 	classlist 	classList �  � � � X   " = ��� � � s   2 8 � � � n   2 5 � � � m   3 5��
�� 
pcls � o   2 3���� 0 selecteditem selectedItem � n       � � �  ;   6 7 � o   5 6���� 0 	classlist 	classList�� 0 selecteditem selectedItem � o   % &���� 0 selecteditems selectedItems �  ��� � Z   > R � ��� � � E   > A � � � o   > ?���� 0 	classlist 	classList � m   ? @��
�� 
cTsk � r   D G � � � m   D E � � � � �  T a s k � o      ���� 0 	the_class  ��   � r   J R � � � l  J P ����� � n   J P � � � m   N P��
�� 
pcls � n   J N � � � 4   K N�� �
�� 
cobj � m   L M����  � o   J K���� 0 selecteditems selectedItems��  ��   � o      ���� 0 	raw_class 	raw_Class��  ��  ��   �  � � � Z  W d � ����� � =  W Z � � � o   W X���� 0 	raw_class 	raw_Class � m   X Y��
�� 
cEvt � r   ] ` � � � m   ] ^ � � � � �  C a l e n d a r � o      ���� 0 	the_class  ��  ��   �  � � � Z  e r � ����� � =  e h � � � o   e f���� 0 	raw_class 	raw_Class � m   f g��
�� 
cNot � r   k n � � � m   k l � � � � �  N o t e � o      ���� 0 	the_class  ��  ��   �  � � � Z  s � � ����� � =  s v � � � o   s t���� 0 	raw_class 	raw_Class � m   t u��
�� 
cTsk � r   y | � � � m   y z � � � � �  T a s k � o      ���� 0 	the_class  ��  ��   �  � � � Z  � � � ����� � =  � � � � � o   � ����� 0 	raw_class 	raw_Class � m   � ��
� 
cAbE � r   � � � � � m   � � � � � � �  C o n t a c t � o      �~�~ 0 	the_class  ��  ��   �  � � � Z  � � � ��}�| � =  � � � � � o   � ��{�{ 0 	raw_class 	raw_Class � m   � ��z
�z 
inm  � r   � � � � � m   � � � � � � �  M e s s a g e � o      �y�y 0 	the_class  �}  �|   �  � � � Z  � � � ��x�w � =  � � � � � o   � ��v�v 0 	raw_class 	raw_Class � m   � ��u
�u 
ctxt � r   � � � � � m   � � � � � � �  T e x t � o      �t�t 0 	the_class  �x  �w   �  ��s � Z  � � � ��r�q � =  � � � � � o   � ��p�p 0 
defaulttag 
defaultTag � m   � � � � � � �   � r   � � � � � o   � ��o�o 0 	the_class   � o      �n�n 0 
defaulttag 
defaultTag�r  �q  �s   }   GET MESSAGES    ~ � � �    G E T   M E S S A G E S { R      �m�l�k
�m .ascrerr ****      � ****�l  �k  ��   y  ��j � L   � � � � o   � ��i�i 0 selecteditems selectedItems�j   v 5     �h ��g
�h 
capp � m     � � � � � * c o m . m i c r o s o f t . O u t l o o k
�g kfrmID  ��   k  ��f � l     �e�d�c�e  �d  �c  �f       �b � � �b   � �a�`�a 0 
item_check 
item_Check
�` .aevtoappnull  �   � **** � �_ m�^�]�\�_ 0 
item_check 
item_Check�^  �]   �[�Z�Y�X�W�V�[ 0 selecteditems selectedItems�Z 0 	raw_class 	raw_Class�Y 0 	classlist 	classList�X 0 selecteditem selectedItem�W 0 	the_class  �V 0 
defaulttag 
defaultTag �U ��T�S�R�Q�P�O�N�M ��L ��K � ��J ��I ��H � ��G�F
�U 
capp
�T kfrmID  
�S 
sele
�R 
pcls
�Q 
list
�P 
kocl
�O 
cobj
�N .corecnte****       ****
�M 
cTsk
�L 
cEvt
�K 
cNot
�J 
cAbE
�I 
inm 
�H 
ctxt�G  �F  �\ �)���0 � �*�,E�O��,E�O��  :jvE�O �[��l kh ��,�6G[OY��O�� �E�Y 
��k/�,E�Y hO��  �E�Y hO��  �E�Y hO��  �E�Y hO�a   
a E�Y hO�a   
a E�Y hO�a   
a E�Y hO�a   �E�Y hW X  hO�U  �E�D�C�B
�E .aevtoappnull  �   � **** k     a  �A�A  �D  �C      `�@�?�>�=�<�;�:�9 @�8�7 F�6 J�5�4�3�2�1�0 ^�/�@ 0 
item_check 
item_Check�? $0 selectedmessages selectedMessages
�> .corecnte****       ****�= 0 msgcount msgCount
�< 
cobj�; "0 selectedmessage selectedMessage
�: 
pcls
�9 
cEvt
�8 
disp
�7 
btns
�6 
dflt
�5 
givu�4 
�3 .sysodlogaskr        TEXT
�2 
ID  
�1 
TEXT�0 0 msgid msgID
�/ .JonspClpnull���     ****�B b� ^)j+ E�O�j E�O�k hY hO��k/E�O��,�  ��l�����la  OhY hO�a ,a &E` Oa _ %j Uascr  ��ޭ