﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("ProjectReviewTool.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to /* 
        '''	Generic Styling, for Desktops/Laptops 
        '''	*/
        '''	body,th,td{ font-family:Arial; font-size:10px;}
        '''	table { 
        '''		width: 100%; 
        '''		border-collapse: collapse;    
        '''	}
        '''	
        '''	
        '''	/* Zebra striping */
        '''	tr:nth-of-type(odd) { 
        '''		background: #eee; 
        '''	}
        '''	th { 
        '''		background: #333; 
        '''		color: white; 
        '''		font-weight: bold; 
        '''	}
        '''	td, th { 
        '''		padding: 6px; 
        '''		border: 1px solid #ccc; 
        '''		text-align: left; 
        '''	}
        '''	tr.selected { 
        '''		background: #700000;
        '''		font-weight:bold;
        '''		color:white;
        '''	}
        '''
        '''   /* 
        '''	Max width bef [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property TableCSS() As String
            Get
                Return ResourceManager.GetString("TableCSS", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;script type=&quot;text/javascript&quot;&gt;
        '''&lt;!--
        '''    function toggle_visibility(id) {
        '''       var e = document.getElementById(id);
        '''       if(e.style.display == &apos;block&apos;)
        '''          e.style.display = &apos;none&apos;;
        '''       else
        '''          e.style.display = &apos;block&apos;;
        '''    }
        '''//--&gt;
        '''&lt;/script&gt;.
        '''</summary>
        Friend ReadOnly Property TableJS() As String
            Get
                Return ResourceManager.GetString("TableJS", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
