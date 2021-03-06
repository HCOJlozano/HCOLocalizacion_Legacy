//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace T1.DBManager.Resources {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class DBManagerRes {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal DBManagerRes() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("T1.DBManager.Resources.DBManagerRes", typeof(DBManagerRes).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Create DATABASE [T1].
        /// </summary>
        internal static string CreateDBScriptSQL {
            get {
                return ResourceManager.GetString("CreateDBScriptSQL", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to CREATE TABLE [dbo].[OBJECTCONTROL](
        ///	[ID] [int] IDENTITY(1,1) NOT NULL,
        ///	[OBJECTDEFID] [int] NOT NULL,
        ///	[LASTREAD] [nvarchar](50) NULL,
        ///	[LASTCORRECT] [nvarchar](50) NULL,
        ///	[LASTREADDATE] [datetime] NULL,
        ///	[LASTCORRECTDATE] [datetime] NULL,
        /// CONSTRAINT [PK_OBJECTCONTROL] PRIMARY KEY CLUSTERED 
        ///(
        ///	[ID] ASC
        ///)
        ///) ON [PRIMARY]
        ///
        ///
        ///.
        /// </summary>
        internal static string CreateTableObjectControlSql {
            get {
                return ResourceManager.GetString("CreateTableObjectControlSql", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to CREATE TABLE [dbo].[OBJECTCRON](
        ///	[ID] [int] IDENTITY(1,1) NOT NULL,
        ///	[OBJECTDEFID] [int] NOT NULL,
        ///	[CRON] [nvarchar](50) NOT NULL,
        /// CONSTRAINT [PK_OBJECTCRON] PRIMARY KEY CLUSTERED 
        ///(
        ///	[ID] ASC
        ///)
        ///) ON [PRIMARY]
        ///
        ///
        ///.
        /// </summary>
        internal static string CreateTableOBJECTCRONSql {
            get {
                return ResourceManager.GetString("CreateTableOBJECTCRONSql", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to CREATE TABLE [dbo].[OBJECTDEF](
        ///	[ID] [int] IDENTITY(1,1) NOT NULL,
        ///	[B1OBJECT] [nvarchar](50) NOT NULL,
        ///	[TABLES] [nvarchar](255) NOT NULL,
        /// CONSTRAINT [PK_OBJECTDEF] PRIMARY KEY CLUSTERED 
        ///(
        ///	[ID] ASC
        ///)
        ///) ON [PRIMARY]
        ///
        ///
        ///.
        /// </summary>
        internal static string CreateTableOBJECTDEFSql {
            get {
                return ResourceManager.GetString("CreateTableOBJECTDEFSql", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to CREATE TABLE [dbo].[OBJECTLOG](
        ///	[ID] [int] IDENTITY(1,1) NOT NULL,
        ///	[OBJECTCONTROLID] [int] NOT NULL,
        ///	[MESSAGE] [nvarchar](max) NULL,
        ///	[LOGDATE] [datetime] NULL,
        ///	[OPERATION] [nvarchar](50) NULL,
        /// CONSTRAINT [PK_OBJECTLOG] PRIMARY KEY CLUSTERED 
        ///(
        ///	[ID] ASC
        ///)
        ///) ON [PRIMARY] 
        ///
        ///
        ///.
        /// </summary>
        internal static string CreateTableOBJECTLOGSql {
            get {
                return ResourceManager.GetString("CreateTableOBJECTLOGSql", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to CREATE TABLE [dbo].[OBJECTOPERATION](
        ///	[ID] [int] IDENTITY(1,1) NOT NULL,
        ///	[OBJECTDEFID] [int] NOT NULL,
        ///	[OBJECTKEY] [nvarchar](100) NULL,
        ///	[OBJECTXML] [nvarchar](max) NULL,
        ///	[OBJECTDATE] [datetime] NULL,
        ///	[STATUS] [nvarchar](50) NULL,
        /// CONSTRAINT [PK_OBJECTIN] PRIMARY KEY CLUSTERED 
        ///(
        ///	[ID] ASC
        ///)
        ///) ON [PRIMARY] 
        ///
        ///
        ///.
        /// </summary>
        internal static string CreateTableOBJECTOPERATIONSql {
            get {
                return ResourceManager.GetString("CreateTableOBJECTOPERATIONSql", resourceCulture);
            }
        }
    }
}
