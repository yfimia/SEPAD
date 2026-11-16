CREATE TABLE [AgendaEventos] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[OwnerName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OwnerId] [int] NULL ,
	[ToName] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL CONSTRAINT [DF_AgendaEventos_ToName] DEFAULT (''),
	[ToUsr] [int] NULL CONSTRAINT [DF_AgendaEventos_ToUsr] DEFAULT (0),
	[ToModulo] [int] NULL CONSTRAINT [DF_AgendaEventos_ToModulo] DEFAULT (0),
	[ToCourse] [int] NULL CONSTRAINT [DF_AgendaEventos_ToCourse] DEFAULT (0),
	[ToGroup] [int] NULL CONSTRAINT [DF_AgendaEventos_ToGroup] DEFAULT (0),
	[ToSubGroup] [int] NULL CONSTRAINT [DF_AgendaEventos_ToSubGroup] DEFAULT (0),
	[Subject] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Content] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Priority] [int] NULL CONSTRAINT [DF_EventosAgenda_Priority] DEFAULT (0),
	[DateBegin] [datetime] NULL ,
	[DateEnd] [datetime] NULL ,
	[EventType] [int] NULL CONSTRAINT [DF_EventosAgenda_EventType] DEFAULT (1),
	[Readed] [bit] NOT NULL CONSTRAINT [DF_EventosAgenda_Readed] DEFAULT (1),
	CONSTRAINT [PK_AgendaEventos] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] ,
	CONSTRAINT [FK_AgendaEventos_Cursos] FOREIGN KEY 
	(
		[ToCourse]
	) REFERENCES [Cursos] (
		[ID]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_AgendaEventos_Grupos] FOREIGN KEY 
	(
		[ToGroup]
	) REFERENCES [Grupos] (
		[ID]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_AgendaEventos_Modulo] FOREIGN KEY 
	(
		[ToModulo]
	) REFERENCES [Modulo] (
		[ID]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_AgendaEventos_SubGrupo] FOREIGN KEY 
	(
		[ToSubGroup]
	) REFERENCES [SubGrupo] (
		[ID]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_AgendaEventos_Usuarios] FOREIGN KEY 
	(
		[ToUsr]
	) REFERENCES [Usuarios] (
		[ID]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_AgendaEventos_Usuarios1] FOREIGN KEY 
	(
		[OwnerId]
	) REFERENCES [Usuarios] (
		[ID]
	) NOT FOR REPLICATION 
) ON [PRIMARY]
GO


alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_Cursos]
GO

alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_Grupos]
GO

alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_Modulo]
GO

alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_SubGrupo]
GO

alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_Usuarios]
GO

alter table [dbo].[AgendaEventos] nocheck constraint [FK_AgendaEventos_Usuarios1]
GO

