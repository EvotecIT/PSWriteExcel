Clear-Host
Import-Module "C:\Support\GitHub\PSPublishModule\PSPublishModule.psm1" -Force

$Configuration = @{
    Information = @{
        ModuleName        = 'PSWriteExcel'
        DirectoryProjects = 'C:\Support\GitHub'
        LibrariesCore     = 'Lib\Core'
        LibrariesDefault  = 'Lib\Default'

        Manifest          = @{
            # Minimum version of the Windows PowerShell engine required by this module
            PowerShellVersion    = '5.1'
            # Supported PSEditions
            CompatiblePSEditions = @('Desktop', 'Core')
            # ID used to uniquely identify this module
            GUID                 = '82232c6a-27f1-435d-a496-929f7221334b'
            # Version number of this module.
            ModuleVersion        = '0.1.X'
            # ID used to uniquely identify this module
            # Author of this module
            Author               = 'Przemyslaw Klys'
            # Company or vendor of this module
            CompanyName          = 'Evotec'
            # Copyright statement for this module
            Copyright            = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
            # Description of the functionality provided by this module
            Description          = 'Little project to create Excel files without Microsoft Excel being installed.'
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags                 = @('Excel', 'ConvertTo-Excel', 'ExportExcel', 'macOS', 'linux', 'windows')
            # A URL to the main website for this project.
            ProjectUri           = 'https://github.com/EvotecIT/PSWriteExcel'

            IconUri              = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteExcel.png'

            RequiredModules      = @(
                @{ ModuleName = 'PSSharedGoods'; ModuleVersion = "Latest"; Guid = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe' }
            )
        }
    }
    Options     = @{
        Merge             = @{
            Sort           = 'None'
            FormatCodePSM1 = @{
                Enabled           = $true
                RemoveComments    = $true
                FormatterSettings = @{
                    IncludeRules = @(
                        'PSPlaceOpenBrace',
                        'PSPlaceCloseBrace',
                        'PSUseConsistentWhitespace',
                        'PSUseConsistentIndentation',
                        'PSAlignAssignmentStatement',
                        'PSUseCorrectCasing'
                    )

                    Rules        = @{
                        PSPlaceOpenBrace           = @{
                            Enable             = $true
                            OnSameLine         = $true
                            NewLineAfter       = $true
                            IgnoreOneLineBlock = $true
                        }

                        PSPlaceCloseBrace          = @{
                            Enable             = $true
                            NewLineAfter       = $false
                            IgnoreOneLineBlock = $true
                            NoEmptyLineBefore  = $false
                        }

                        PSUseConsistentIndentation = @{
                            Enable              = $true
                            Kind                = 'space'
                            PipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
                            IndentationSize     = 4
                        }

                        PSUseConsistentWhitespace  = @{
                            Enable          = $true
                            CheckInnerBrace = $true
                            CheckOpenBrace  = $true
                            CheckOpenParen  = $true
                            CheckOperator   = $true
                            CheckPipe       = $true
                            CheckSeparator  = $true
                        }

                        PSAlignAssignmentStatement = @{
                            Enable         = $true
                            CheckHashtable = $true
                        }

                        PSUseCorrectCasing         = @{
                            Enable = $true
                        }
                    }
                }
            }
            FormatCodePSD1 = @{
                Enabled        = $true
                RemoveComments = $false
            }
            Integrate      = @{
                ApprovedModules = @('PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword')
            }
        }
        Standard          = @{
            FormatCodePSM1 = @{

            }
            FormatCodePSD1 = @{
                Enabled = $true
                #RemoveComments = $true
            }
        }
        PowerShellGallery = @{
            ApiKey   = 'C:\Support\Important\PowerShellGalleryAPI.txt'
            FromFile = $true
        }
        GitHub            = @{
            ApiKey   = 'C:\Support\Important\GithubAPI.txt'
            FromFile = $true
            UserName = 'EvotecIT'
            #RepositoryName = 'PSWriteHTML'
        }
        Documentation     = @{
            Path       = 'Docs'
            PathReadme = 'Docs\Readme.md'
        }
    }
    Steps       = @{
        BuildModule        = @{  # requires Enable to be on to process all of that
            Enable              = $true
            DeleteBefore        = $true
            Merge               = $true
            MergeMissing        = $true
            LibrarySeparateFile = $true
            LibraryDotSource    = $true
            ClassesDotSource    = $false
            SignMerged          = $true
            CreateFileCatalog   = $false # not working
            Releases            = $true
            ReleasesUnpacked    = $false
            RefreshPSD1Only     = $false
        }
        BuildDocumentation = $true
        ImportModules      = @{
            Self            = $true
            RequiredModules = $false
            Verbose         = $false
        }
        PublishModule      = @{  # requires Enable to be on to process all of that
            Enabled      = $true
            Prerelease   = ''
            RequireForce = $false
            GitHub       = $true
        }
    }
}

New-PrepareModule -Configuration $Configuration