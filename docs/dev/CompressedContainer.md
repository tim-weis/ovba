# CompressedContainer

The published documentation ([\[MS-OVBA\]: Office VBA File Format Structure](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc)) explains how to decode a CompressedContainer. The algorithm descibed is rather complex, and appears to have been derived from an existing implementation (as opposed to the implementation following the specification).

As it turns out, the implementation can be simplified a lot, once the essence of the problem has been crystallized. The following text describes how to arrive at an almost trivial implementation.
