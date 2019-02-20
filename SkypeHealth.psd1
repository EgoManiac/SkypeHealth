#
# Module manifest for module 'SkypeHealth'
#
# Generated by: Chris Burns
#
# Generated on: 18/04/2017
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule        = 'SkypeHealth.psm1'
    
    # Version number of this module.
    ModuleVersion     = '2.2'
    
    # ID used to uniquely identify this module
    GUID              = '7d8813e8-6630-4eff-827f-11b121a2d8fc'
    
    # Author of this module
    Author            = 'Chris Burns'
    
    # Company or vendor of this module
    CompanyName       = 'CB'
    
    # Copyright statement for this module
    Copyright         = '(c) 2018 Chris Burns. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description       = 'Welcome to Skype/Lync Health Modules. This module has been designed to provide a very simple report that provides an overview of your environment in a simple HTML file. The outputted HTML file is designed to be portable and to work on iphone/ipad devives so can be emailed to managers/directors.'
    
    # Minimum version of the Windows PowerShell engine required by this module
    PowerShellVersion = '4.0'
    
    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''
    
    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''
    
    # Minimum version of Microsoft .NET Framework required by this module
    # DotNetFrameworkVersion = ''
    
    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''
    
    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''
    
    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()
    
    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()
    
    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()
    
    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()
    
    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()
    
    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()
    
    # Functions to export from this module
    FunctionsToExport = @('Get-CBSkypeHealth', 'Get-CBlyncBase64')
    
    # Cmdlets to export from this module
    CmdletsToExport   = '*'
    
    # Variables to export from this module
    VariablesToExport = '*'
    
    # Aliases to export from this module
    AliasesToExport   = '*'
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
    # ModuleList = @()
    
    # List of all files packaged with this module
    # FileList = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData       = @{
    
        PSData = @{
    
            # Tags applied to this module. These help with module discovery in online galleries.
            # Tags = @()
    
            # A URL to the license for this module.
            # LicenseUri = ''
    
            # A URL to the main website for this project.
            # ProjectUri = ''
    
            # A URL to an icon representing this module.
            # IconUri = ''
    
            # ReleaseNotes of this module
            # ReleaseNotes = ''
    
        } # End of PSData hashtable
    
    } # End of PrivateData hashtable
    
    # HelpInfo URI of this module
    # HelpInfoURI = ''
    
    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''
    
}
    
    