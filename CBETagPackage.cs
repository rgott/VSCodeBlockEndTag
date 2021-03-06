﻿//------------------------------------------------------------------------------
// <copyright file="CBETagPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Utilities;
using System.Collections.Generic;
using System.ComponentModel.Composition;

namespace CodeBlockEndTag
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(CBETagPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    // Add OptionPage to package
    [ProvideOptionPage(typeof(CBEOptionPage), "KC Extensions", "CodeBlock End Tagger", 113, 114, true)]
    // Load package at every (including none) project type
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    public sealed class CBETagPackage : Package
    {
        /// <summary>
        /// CBETagPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "d7c91e0f-240b-4605-9f35-accf63a68623";

        public delegate void PackageOptionChangedHandler(object sender);
        
        /// <summary>
        /// Event fired if any option in the OptionPage is changed
        /// </summary>
        public event PackageOptionChangedHandler PackageOptionChanged;

        public static CBETagPackage Instance => new CBETagPackage();

        private CBETagPackage()
        {
        }

        /// <summary>
        /// Gets a list of all possible content types in VisualStudio
        /// </summary>
        public static IList<IContentType> ContentTypes { get; private set; }

        /// <summary>
        /// Load the list of content types
        /// </summary>
        internal static void ReadContentTypes(IContentTypeRegistryService ContentTypeRegistryService)
        {
            if (ContentTypes != null) return;
            ContentTypes = new List<IContentType>();
            foreach (var ct in ContentTypeRegistryService.ContentTypes)
            {
                if (ct.IsOfType("code"))
                    ContentTypes.Add(ct);
            }
        }

        public static bool IsLanguageSupported(string lang)
        {
            return GetOptions().IsLanguageSupported(lang);
        }

        #region Option Values
        
        public static int CBEDisplayMode
        {
            get
            {
                return GetOptions().CBEDisplayMode;
            }
        }

        public static int CBEVisibilityMode
        {
            get
            {
                return GetOptions().CBEVisibilityMode;
            }
        }

        public static bool CBETaggerEnabled
        {
            get
            {
                return GetOptions().CBETaggerEnabled;
            }
        }

        public static int CBEClickMode
        {
            get
            {
                return GetOptions().CBEClickMode;
            }
        }

        public static double CBETagScale
        {
            get
            {
                return GetOptions().CBETagScale;
            }
        }

               
        private static CBEOptionPage GetOptions()
        {
            return (CBEOptionPage)Instance.GetDialogPage(typeof(CBEOptionPage));
        }

        #endregion

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
            var page = (CBEOptionPage)Instance.GetDialogPage(typeof(CBEOptionPage));
            page.OptionChanged += Page_OptionChanged;
        }

        private void Page_OptionChanged(object sender)
        {
            PackageOptionChanged?.Invoke(this);
        }

        #endregion
    }
}
