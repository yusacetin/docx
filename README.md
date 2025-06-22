## [WIP] Simple Office Open XML Document (docx) Library

A simple library for creating (very) simple docx files, made with [my XML library](https://github.com/yusacetin/xml). Same philosophy. Requires the zip program to be installed and only runs on POSIX systems (due to system call, might make cross platform later).

### Structure

The library consists of a single hpp file (`docx.hpp`) that has three classes: `DOCX`, `Paragraph`, and `Text`. A `DOCX` object is basically a vector of `Paragraph`s, and a `Paragraph` is basically a vector of `Text`s. `DOCX` and `Paragraph` classes have `get()` methods that return an `XML::Node` object. When the `get()` method of a `DOCX` object is called, it iterates through its vector of `Paragraph`s and calls each of their `get()` methods, which in turn generate the `XML::Node` object corresponding to that `Paragraph`, and adds the returned `XML::Node` objects together to get the complete XML file. The docx zip file is generated using the `zip` program via a system call. See `main.cpp` for a usage example.

### License

GNU General Public License version 3 or later.