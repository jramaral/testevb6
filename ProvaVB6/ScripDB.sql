
USE master

CREATE DATABASE CountryDb
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'CountryDb', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL12.SQLEXPRESS\MSSQL\DATA\CountryDb.mdf' , SIZE = 5120KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'CountryDb_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL12.SQLEXPRESS\MSSQL\DATA\CountryDb_log.ldf' , SIZE = 2048KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)



Use CountryDb

create table CountryInfo(
	Cod int NOT NULL,
	sISOCode nchar(3) NULL,
	sName nchar(50) NULL,
	sCapitalCity nchar(50) NULL,
	sPhoneCode int NULL,
	sContinentCode nchar(3) NULL,
	sCurrencyISOCode nchar(4) NULL,
	sCountryFlag nchar(150) NULL,
	primary key (Cod)

)


create table Languages (
  IsoCode varchar(4) null,
  sNameLang varchar(50) null,
  Cod_Country int not null
  
)

alter table Languages add constraint lang_fk foreign key ( Cod_Country ) references CountryInfo( Cod );