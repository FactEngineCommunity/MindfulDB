# MindfulDB
Graph Relational DBMS

<p align="center">
  <img src="https://github.com/FactEngineCommunity/MindfulDB/assets/10895608/42af2d95-83b7-4afb-86b1-8d11c6cfc9c3" />
</p>

# Who | What is MindfulDB?

MindfulDB = Cypher over your otherwise relational database.

MindfulDB turns your otherwise relational database into a **graph relational database**.

MindfulDB is realisation of graph relational paradigm by extending your existing relational database to act as if a graph database from a query perspective.

Without changing anything about your relational schema MindfulDB empowers you to perform Cypher queries over your database with the minimum of fuss.

SQL + Cypher on your MindfulDB.

# How does MindfulDB Work?

MindfulDB simply uses EDGE LABELS injected into the comment section of your otherwise relational database schema so that Cypher queries can be run over your otherwise relational database.

# What are the steps involved to get MindfulDB up and running?

1. Download a copy of Boston from www.FactEngine.ai;
2. Add the EDGE LABELS for all Edge Types in the schema and save it
      Reverse engineer the schema of your otherwise relational database in Boston to get a Property Graph Schema and a Entity Relationship Diagram of your schema;
      Save the EDGE LABELS to the comments on your otherwise relational database schema (as JSON in the comments);
3. Download and start using the MindfulDB DLLs (Direct Link Libraries) to write Cypher queries against your now graph relational database.

# Software Languages/Platforms Supported

Any software language/platform that supports DLLs (Direct Link Libraries) at this stage;

# Database Supported

SQLite at this stage

