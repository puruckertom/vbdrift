VERSION 5.00
Begin VB.Form frmLicense 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AgDRIFT End-User License Agreement"
   ClientHeight    =   6345
   ClientLeft      =   2985
   ClientTop       =   1680
   ClientWidth     =   7530
   ForeColor       =   &H80000008&
   HelpContextID   =   1455
   Icon            =   "LICENSE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   7530
   Begin VB.TextBox Text1 
      Height          =   4455
      HelpContextID   =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1320
      Width           =   7335
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2280
      ScaleHeight     =   735
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.Label lblTM 
         AutoSize        =   -1  'True
         Caption         =   "�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   195
      End
      Begin VB.Line linLogo 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         X1              =   720
         X2              =   2400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblLogo 
         Caption         =   "AgDRIFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1455
      Left            =   3360
      TabIndex        =   0
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Spray Drift Task Force Spray Software"
      Height          =   195
      Left            =   2190
      TabIndex        =   1
      Top             =   960
      Width           =   2745
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: license.frm,v 1.4 2001/05/24 20:16:21 tom Exp $

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Static s As String
  
  CenterForm Me

  s = ""
  s = s + "END-USER LICENSE AGREEMENT FOR SPRAY DRIFT TASK FORCE SOFTWARE" + vbCrLf
  s = s + "" + vbCrLf
  s = s + "IMPORTANT-READ CAREFULLY: This End-User License Agreement (�LICENSE " + vbCrLf
  s = s + "AGREEMENT�) is a legally binding agreement between you (whether an individual " + vbCrLf
  s = s + "and/or an entity) as the �Licensee� and the Spray Drift Task Force (�Task Force�) as the " + vbCrLf
  s = s + "�Licensor� concerning installation and use of the AgDRIFT� model, which includes the " + vbCrLf
  s = s + "computer software release specified therein, associated media, printed materials, " + vbCrLf
  s = s + "incorporated data, �online� or electronic documentation, and output generated by or from " + vbCrLf
  s = s + "the AgDRIFT� model (collectively, �SOFTWARE PRODUCT�). Your installation and use " + vbCrLf
  s = s + "of the SOFTWARE PRODUCT is subject to the terms and conditions of this LICENSE " + vbCrLf
  s = s + "AGREEMENT and all applicable laws." + vbCrLf
  s = s + "  " + vbCrLf
  s = s + "By installing, copying, or otherwise using the SOFTWARE PRODUCT, you agree to be " + vbCrLf
  s = s + "bound by the terms of this LICENSE AGREEMENT. If you do not agree to the terms of " + vbCrLf
  s = s + "this LICENSE AGREEMENT, do not install or use the SOFTWARE PRODUCT and " + vbCrLf
  s = s + "delete the SOFTWARE PRODUCT from your computer or return it to the Task Force." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "SOFTWARE PRODUCT LICENSE" + vbCrLf
  s = s + "" + vbCrLf
  s = s + "The SOFTWARE PRODUCT is protected by copyright laws and international copyright " + vbCrLf
  s = s + "treaties, as well as other intellectual property laws and treaties. The rights granted to " + vbCrLf
  s = s + "you by this LICENSE AGREEMENT constitute a license, not a transfer of title." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "1. GRANT OF LICENSE. Except as otherwise provided herein, this LICENSE " + vbCrLf
  s = s + "AGREEMENT grants you a limited, royalty-free, revocable and non-transferable license " + vbCrLf
  s = s + "for the right to use the SOFTWARE PRODUCT including the object code of the " + vbCrLf
  s = s + "software release specified therein during the term hereof as an educational, " + vbCrLf
  s = s + "governmental, and management tool for assessing and mitigating pesticide spray drift " + vbCrLf
  s = s + "and for educating applicators of pesticides in pesticide spray drift reduction. As a " + vbCrLf
  s = s + "Licensee, you shall not decompile, reverse engineer or otherwise attempt to discover " + vbCrLf
  s = s + "the source code of the SOFTWARE PRODUCT nor authorize or permit others to do so. " + vbCrLf
  s = s + "Non-governmental Licensees shall not use the SOFTWARE PRODUCT in any litigation, " + vbCrLf
  s = s + "including without limitation judicial and administrative proceedings. To install or use the " + vbCrLf
  s = s + "SOFTWARE PRODUCT in a manner not expressly licensed by the terms and " + vbCrLf
  s = s + "conditions of this LICENSE AGREEMENT, you must seek a modification of the " + vbCrLf
  s = s + "LICENSE AGREEMENT as provided for herein, which modification the Task Force in its " + vbCrLf
  s = s + "sole discretion may agree to provide." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "2. PESTICIDE REGISTRATION DATA. Notwithstanding the foregoing GRANT OF " + vbCrLf
  s = s + "LICENSE, the SOFTWARE PRODUCT shall not be used to support directly or indirectly " + vbCrLf
  s = s + "the registration, continued registration, or amended registration of a pesticide product - " + vbCrLf
  s = s + "including without limitation pesticide registration, reregistration, tolerance exemption, " + vbCrLf
  s = s + "tolerance setting, and any related risk or exposure assessments, whether in the United " + vbCrLf
  s = s + "States or any of its states, territories, or other subdivisions or in any jurisdiction outside " + vbCrLf
  s = s + "the United States - to benefit a person or entity unless such person or entity is " + vbCrLf
  s = s + "expressly authorized in writing by the Task Force to rely upon Task Force data. In " + vbCrLf
  s = s + "addition to the foregoing, governmental entities performing the above-described " + vbCrLf
  s = s + "pesticide-related regulatory purposes may use the SOFTWARE PRODUCT for such " + vbCrLf
  s = s + "pesticide-related regulatory purposes for the benefit of a person or entity entitled to rely " + vbCrLf
  s = s + "upon Task Force data by virtue of a binding offer to pay compensation for such use " + vbCrLf
  s = s + "under � 3(c)(1)(F)(iii) of the Federal Insecticide, Fungicide, and Rodenticide Act " + vbCrLf
  s = s + "(FIFRA), 7 U.S.C. � 136a(c)(1)(F)(iii), or other provisions of law that provide identical " + vbCrLf
  s = s + "protections to owners and submitters of data. You recognize that the SOFTWARE " + vbCrLf
  s = s + "PRODUCT is based upon, and validated for the foregoing pesticide-related regulatory " + vbCrLf
  s = s + "purposes by data owned by the Task Force." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "3. INTELLECTUAL PROPERTY; CONFIDENTIALITY. The SOFTWARE PRODUCT is " + vbCrLf
  s = s + "the intellectual property of the Task Force and is licensed to each Licensee only as " + vbCrLf
  s = s + "expressly provided by this LICENSE AGREEMENT. All other rights are reserved. The " + vbCrLf
  s = s + "Task Force�s intellectual property rights in the SOFTWARE PRODUCT include, without " + vbCrLf
  s = s + "limitation, all data incorporated into and all output generated by or from the SOFTWARE " + vbCrLf
  s = s + "PRODUCT and each reference herein to the SOFTWARE PRODUCT includes, without " + vbCrLf
  s = s + "limitation, such data and output. You shall not remove any copyright or other proprietary " + vbCrLf
  s = s + "notices contained in the SOFTWARE PRODUCT and may not distribute, transfer or " + vbCrLf
  s = s + "publish the SOFTWARE PRODUCT without the express written permission of the Task " + vbCrLf
  s = s + "Force. AgDRIFT� is a registered trademark of the Spray Drift Task Force. Each end " + vbCrLf
  s = s + "user is responsible for maintaining the confidentiality of their own account number " + vbCrLf
  s = s + "and/or password, if applicable." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "4. GOVERNING LAW AND JURISDICTION. By installing or using this SOFTWARE " + vbCrLf
  s = s + "PRODUCT, you agree that all matters relating to such installation or use shall be " + vbCrLf
  s = s + "governed by the statutes and laws of the District of Columbia, without regard to the " + vbCrLf
  s = s + "conflicts of law principles thereof. You and the Task Force agree and hereby submit to " + vbCrLf
  s = s + "the personal jurisdiction of, and exclusive venue in, the Superior Court of District of " + vbCrLf
  s = s + "Columbia and the United States District Court for the District of Columbia with respect to " + vbCrLf
  s = s + "such matters, except that - for matters regarding the amount of compensation owed to " + vbCrLf
  s = s + "the Task Force arising from use of the SOFTWARE PRODUCT - you or the Task Force " + vbCrLf
  s = s + "may initiate binding arbitration in the District of Columbia under the terms of this " + vbCrLf
  s = s + "LICENSE AGREEMENT. Notwithstanding the foregoing, in the event that Task Force " + vbCrLf
  s = s + "finds, in its sole discretion, that it might suffer irreparable loss or harm to any of its " + vbCrLf
  s = s + "intellectual property rights as a result of your breach of this LICENSE AGREEMENT, " + vbCrLf
  s = s + "then the Task Force may, but shall not be required to, pursue any and all equitable or " + vbCrLf
  s = s + "legal remedies in any court of competent jurisdiction pending outcome of the arbitration " + vbCrLf
  s = s + "proceedings as otherwise provided for herein." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "5. ARBITRATION OF COMPENSATION CLAIMS. If the Task Force loses any data " + vbCrLf
  s = s + "compensation rights in the SOFTWARE PRODUCT as the result of breach of this " + vbCrLf
  s = s + "LICENSE AGREEMENT by you or any person or entity for which you are responsible " + vbCrLf
  s = s + "under this LICENSE AGREEMENT, you offer and agree to compensate the Task Force " + vbCrLf
  s = s + "for such loss (including attorneys fees) and agree to submit such matter to binding " + vbCrLf
  s = s + "arbitration in the same manner provided by FIFRA � 3(c)(1)(F)(iii), 7 U.S.C. " + vbCrLf
  s = s + "� 136a(c)(1)(F)(iii)." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "6. INDEMNIFICATION. You are responsible for any and all use of the SOFTWARE " + vbCrLf
  s = s + "PRODUCT that you install, and you agree to indemnify, defend, and hold harmless the " + vbCrLf
  s = s + "Task Force from any liability or expense arising from such use or misuse. You further " + vbCrLf
  s = s + "agree to immediately notify the Task Force of any unauthorized use of the SOFTWARE " + vbCrLf
  s = s + "PRODUCT that you install or any other breach of this LICENSE AGREEMENT known to " + vbCrLf
  s = s + "you." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "7. U.S. GOVERNMENT RESTRICTED RIGHTS. The Cooperative Research and " + vbCrLf
  s = s + "Development Agreement (CRADA) between the U.S. Environmental Protection Agency " + vbCrLf
  s = s + "(EPA) Office of Research and Development, the U.S. Department of Agriculture (USDA) " + vbCrLf
  s = s + "Agricultural Research Service (ARS), USDA Forest Service, and the Task Force " + vbCrLf
  s = s + "establishes certain rights and duties with respect to the SOFTWARE PRODUCT, " + vbCrLf
  s = s + "including a nonexclusive, irrevocable, paid-up, worldwide Government-Purpose license " + vbCrLf
  s = s + "for the U.S. Government, the terms of which are incorporated herein by reference. " + vbCrLf
  s = s + "FIFRA and the Federal Food, Drug and Cosmetic Act (�FFDCA�) establish certain U.S. " + vbCrLf
  s = s + "Government rights and duties with respect to data, such as the SOFTWARE " + vbCrLf
  s = s + "PRODUCT, submitted in support of certain pesticide-related regulatory actions.  This " + vbCrLf
  s = s + "LICENSE AGREEMENT does not affect or alter U.S. Government rights and duties " + vbCrLf
  s = s + "under the CRADA, FIFRA, or FFDCA. In addition to the foregoing rights, the U.S. " + vbCrLf
  s = s + "Government may distribute or disclose the SOFTWARE PRODUCT within the U.S. " + vbCrLf
  s = s + "Government pursuant to subparagraphs (c)(1) and (2) of the Commercial Computer " + vbCrLf
  s = s + "Software - Restricted Rights at 48 CFR � 52.227-19, provided that the rights granted " + vbCrLf
  s = s + "hereby under such subparagraphs shall not extend to distribution or disclosure in a " + vbCrLf
  s = s + "manner - including without limitation posting on an Internet host server or Internet site - " + vbCrLf
  s = s + "accessible to persons not expressly included under such subparagraphs. Except as " + vbCrLf
  s = s + "otherwise expressly provided by the CRADA, FIFRA, FFDCA, or this LICENSE " + vbCrLf
  s = s + "AGREEMENT, the SOFTWARE PRODUCT is licensed to U.S. Government end users " + vbCrLf
  s = s + "with only those rights granted to all other end users pursuant to the terms and " + vbCrLf
  s = s + "conditions herein." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "8. TASK FORCE MEMBER RIGHTS. The Task Force Joint Data Development " + vbCrLf
  s = s + "Agreement establishes certain rights and duties with respect to the SOFTWARE " + vbCrLf
  s = s + "PRODUCT on the part of Task Force members and certain affiliated entities. This " + vbCrLf
  s = s + "LICENSE AGREEMENT does not affect or alter such rights and duties." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "9. OTHER THIRD-PARTY RIGHTS. This SOFTWARE PRODUCT is licensed subject to " + vbCrLf
  s = s + "all intellectual property ownership and other rights held by the Task Force's licensors, " + vbCrLf
  s = s + "suppliers, and other providers, including without limitation Task Force members, " + vbCrLf
  s = s + "government agencies, and software vendors. No rights or duties under this LICENSE " + vbCrLf
  s = s + "AGREEMENT affect or alter such third-party ownership or rights." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "10. EXPORT RESTRICTIONS. You agree that you do not intend to and will not, directly " + vbCrLf
  s = s + "or indirectly, export, transmit, or permit the exportation or transmission of the " + vbCrLf
  s = s + "SOFTWARE PRODUCT to any country to which or person to whom such export or " + vbCrLf
  s = s + "transmission is restricted by any applicable U.S. regulation or statute, without the prior " + vbCrLf
  s = s + "written consent, if required, of the Bureau of Export Administration of the U.S. " + vbCrLf
  s = s + "Department of Commerce, or such other governmental entity as may have jurisdiction " + vbCrLf
  s = s + "over such export or transmission. You agree that you do not intend to and will not, " + vbCrLf
  s = s + "directly or indirectly, post or permit the posting of the SOFTWARE PRODUCT on an " + vbCrLf
  s = s + "Internet host server or Internet site." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "11. TERM AND TERMINATION. This LICENSE AGREEMENT enters into effect " + vbCrLf
  s = s + "immediately and automatically upon your first installation or use of the SOFTWARE " + vbCrLf
  s = s + "PRODUCT and continues in effect until the Task Force discontinues the release " + vbCrLf
  s = s + "licensed hereunder; PROVIDED, however, that your obligations as a Licensee herein " + vbCrLf
  s = s + "survive the expiration or termination of your license to install or to use the SOFTWARE " + vbCrLf
  s = s + "PRODUCT. Without prejudice to any other rights, Task Force may terminate this " + vbCrLf
  s = s + "LICENSE AGREEMENT if you fail to comply with the terms and conditions of this " + vbCrLf
  s = s + "LICENSE AGREEMENT. In such event, you must destroy all copies of the SOFTWARE " + vbCrLf
  s = s + "PRODUCT." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "12. MODIFICATIONS. This LICENSE AGREEMENT cannot be modified to extend the " + vbCrLf
  s = s + "terms or conditions of the license except by written agreement between you and the " + vbCrLf
  s = s + "Task Force that expressly supercedes or modifies this LICENSE AGREEMENT." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "13. MISCELLANEOUS. Should you have any questions concerning this LICENSE " + vbCrLf
  s = s + "AGREEMENT or desire to license additional rights not granted by this LICENSE " + vbCrLf
  s = s + "AGREEMENT, to publish any information protected by copyright laws and/or this " + vbCrLf
  s = s + "LICENSE AGREEMENT, or to contact Task Force for any reason, please write the " + vbCrLf
  s = s + "Spray Drift Task Force, care of McKenna & Cuneo, L.L.P. / 1900 K Street, NW, Suite " + vbCrLf
  s = s + "100 / Washington, DC 20006-1108." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "14. WARRANTY DISCLAIMER: TASK FORCE (INCLUDING ITS OFFICERS, " + vbCrLf
  s = s + "MEMBERS, AND AGENTS) EXPRESSLY DISCLAIMS ANY WARRANTY FOR THE " + vbCrLf
  s = s + "SOFTWARE PRODUCT. THE SOFTWARE PRODUCT AND ANY RELATED " + vbCrLf
  s = s + "DOCUMENTATION IS PROVIDED �AS IS� WITHOUT WARRANTY OF ANY KIND, " + vbCrLf
  s = s + "EITHER EXPRESS OR IMPLIED, INCLUDING, WITHOUT LIMITATION, THE IMPLIED " + vbCrLf
  s = s + "WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, " + vbCrLf
  s = s + "OR NONINFRINGEMENT OR OF THE QUALITY, ACCURACY, SUITABILITY OF THE " + vbCrLf
  s = s + "RESULTS GENERATED FROM THE USE OF THE SOFTWARE PRODUCT FOR " + vbCrLf
  s = s + "PREDICTING SPRAY DRIFT, WHICH DEPENDS ON ACTUAL USE CONDITIONS. " + vbCrLf
  s = s + "THE ENTIRE RISK ARISING OUT OF USE OR PERFORMANCE OF THE " + vbCrLf
  s = s + "SOFTWARE PRODUCT REMAINS WITH YOU." + vbCrLf
  s = s + "" + vbCrLf
  s = s + "15. NO LIABILITY FOR DAMAGES. In no event shall Task Force (including its Officers, " + vbCrLf
  s = s + "Members, or agents) be liable for any damages whatsoever (including, without " + vbCrLf
  s = s + "limitation, indirect or consequential damages for loss of business profits, business " + vbCrLf
  s = s + "interruption, loss of business information, or any other pecuniary loss) arising out of the " + vbCrLf
  s = s + "use of or inability to use this Task Force product, even if it has been advised of the " + vbCrLf
  s = s + "possibility of such damages. Without limiting the above, the Task Force (including its " + vbCrLf
  s = s + "Officers, Members, or agents) specifically disclaims any responsibility or liability for the " + vbCrLf
  s = s + "reliability or accuracy of the results generated through use of the SOFTWARE " + vbCrLf
  s = s + "PRODUCT to determine or control spray drift or for updating or supporting the " + vbCrLf
  s = s + "SOFTWARE PRODUCT. Because some states/jurisdictions do not allow the exclusion " + vbCrLf
  s = s + "or limitation of liability for consequential or incidental damages, the above limitation may " + vbCrLf
  s = s + "not apply to you."
  Text1.Text = s
End Sub
