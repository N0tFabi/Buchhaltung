<p align="center">
  <img src="https://i.discord.fr/PSS.png">
</p>

<h1 align="center">Buchhaltung</h1>
<p align="center">
  <img src="https://img.shields.io/badge/Discord-%40notfabi-%235464f4">
</p>

If you're too lazy to do it manually

This is a program in c# with which you can do bookkeeping.

Ans also I'll never write a documentation or Unit Tests for this, thanks for asking.

## Usage

### Eröffnungsbilanzkonto (EBK)
Um den Gewinn / Verlust vom letzten Jahr in das neue Jahr zu übertragen
```csharp
Konto ebk = new Konto();
ebk.SetKontoNummer(9800);

ebk.AddSollBetrag(<Betrag : int>, <Datum : string), <Gegenkonten : int[]>);
ebk.AddHabenBetrag(<Betrag : int>, <Datum : string), <Gegenkonten : int[]>);
```

### Buchungssatz
```csharp
Buchungssatz buchungssatz = new Buchungssatz("<sollKonto> <Wert> EUR / <habenKonto> <Wert> EUR");
buchungssatz.SetDatum(<Jahr : int>, <Monat : Monat>, <Tag : int>);
```
Mehrere soll- oder habenKonten:
```csharp
Buchungssatz buchungssatz = new Buchungssatz("<sollKonto_1> <Wert_1> EUR <sollKonto_2> <Wert_2> EUR / <habenKonto_1> <Wert_1> EUR <habenKonto_2> <Wert_2> EUR");
buchungssatz.SetDatum(<Jahr : int>, <Monat : Monat>, <Tag : int>);
```

### Bilanz
#### Neue Bilanz erstellen:
```csharp
Bilanz bilanz = new Bilanz();
```

#### Eröffnungsbilanzkonto (EBK) hinzufügen:
```csharp
bilanz.AddKonto(ebk);
```

#### Bilanz eröffnen (nur nötig, wenn man EBK hinzugefügt hat):
```csharp
bilanz.Open();
```

#### Buchungssatz zur Bilanz hinzufügen:
```csharp
bilanz.AddBuchung(<Buchungssatz : Buchungssatz>);
```

#### Steuern umbuchen (USt und VSt zu UST-Zahllast verbuchen):
```csharp
bilanz.SteuernUmbuchen(<Datum : string>);
```

#### Erfolgskonten mit Warenvorrat laut Inventur abschließen:
```csharp
bilanz.ErfolgsKontenAbschlieszen(<Warenvorrat : decimal>);
```

#### Alle Konten, auf denen etwas verbucht wurde, ausgeben:
```csharp
bilanz.Print();
```

#### Schlussbilanzkonto (SBK) von bestimmten Jahr erstellen und ausgeben:
```csharp
bilanz.CreateSchlussbilanz(<Jahr : int>);
bilanz.PrintSchlussbilanz();
```

#### Bilanz exportieren:
```csharp
bilanz.WriteToExcelFile(<string : filePath>);
```

It is imperative to clarify that you are expressly forbidden from asserting ownership of this program as your exclusive creation. While you may have introduced adjustments or enhancements to it, the fundamental program and its initial design continue to be the intellectual property of the lawful owner or developer. When employing this program, you are obligated to attribute the core program to me as the original creator.
Thanks for your understanding.
