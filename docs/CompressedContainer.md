# CompressedContainer

The published documentation ([\[MS-OVBA\]: Office VBA File Format Structure](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc)) explains how to decode a CompressedContainer. The algorithm descibed is rather complex, and appears to have been derived from an existing implementation, rather than the implementation following the specification.

As it turned out, the implementation can be simplified a lot, once the individual bits of information are put together to form a coherent mental model of the protocol. The following text starts with a high level overview, then gradually adds details and insights, and finally presents an almost trivial implementation of a decompressor.

## Overview

A `CompressedContainer` is a byte stream consisting of a signature byte (`0x01`) followed by a sequence of one or more `CompressedChunk`s. Each `CompressedChunk` starts with a 2-byte `CompressedChunkHeader` that encodes the chunk's size, a signature, and a flag indicating whether the chunk data is compressed or uncompressed<sup>1</sup>.

The header is followed by a stream of bytes (`CompressedChunkData`) that contain either uncompressed data, or an array of `TokenSequence`, controlled by the flag stored in the header. When decompressing, uncompressed data is copied verbatim to the output stream.

A `TokenSequence` starts with a `FlagByte` followed by up to 8 tokens. The individual bits of the `FlagByte` encode the types of tokens following it. There are two token types:

* `LiteralToken` (indicated by `0b0`):

  A `LiteralToken` consists of a single byte. When decompressing, this byte is written to the output stream.
* `CopyToken` (indicated by `0b1`):

  A `CopyToken` is two bytes wide. It encodes offset and length information. The offset is an index into the output stream relative to the current end of that stream. When decompressing, the respective number of bytes, starting at the specified offset, are copied to the end of the output stream<sup>2</sup>.

This concludes the high-level description of the `CompressedContainer`'s layout. The following text contains more detailed information on the individual parts.

---

<sup>1</sup> *A `CompressedChunk` can wind up holding uncompressed data if the compression algorithm has determined that compressing the data doesn't yield any space savings.*

<sup>2</sup> *The tail end of a `CopyToken` can reference data that only becomes available during the process of decompressing it (i.e. length is greater then the end-relative offset). Consider the sequence `AAAA` that can be encoded as a `LiteralToken` (the first `A`) followed by a `CopyToken` with offset 1 and length 3. When decompressing the `CopyToken` the output stream is only one byte long yet the token requests to copy three bytes. This turned out to be a real use case and needed to be addressed in the implementation.*