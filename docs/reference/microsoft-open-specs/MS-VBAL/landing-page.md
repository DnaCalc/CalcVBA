---
layout: Conceptual
title: '[MS-VBAL]: VBA Language Specification | Microsoft Learn'
canonicalUrl: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74
ms.service: openspecs-office
ROBOTS: INDEX, FOLLOW
uhfHeaderId: MSDocsHeader-OpenSpecs
ms.topic: reference
ms.author: cindyle
protocol_rendering: true
description: Specifies the VBA Language, which defines the implementation-independent and operating system-independent programming
locale: en-us
author: mrsgit09
document_id: dc3738ab-58d2-faa6-97bf-4449a4ca86a7
document_version_independent_id: 791ce258-4bf5-b637-ae45-c36ebc3c5516
updated_at: 2025-07-11T20:45:00.0000000Z
original_content_git_url: https://github.com/MicrosoftDocs/open_specs_office/blob/live/documentation/microsoft_general_purpose_programming_languages/MS-VBAL/d5418146-0bd2-45eb-9c7a-fd9502722c74.md
gitcommit: https://github.com/MicrosoftDocs/open_specs_office/blob/88754cb4f6d2de7aeb9cb18c0c83c943d47a0dde/documentation/microsoft_general_purpose_programming_languages/MS-VBAL/d5418146-0bd2-45eb-9c7a-fd9502722c74.md
git_commit_id: 88754cb4f6d2de7aeb9cb18c0c83c943d47a0dde
site_name: Docs
depot_name: MSDN.open_specs_office
page_type: conceptual
toc_rel: toc.json
feedback_system: None
feedback_product_url: ''
feedback_help_link_type: ''
feedback_help_link_url: ''
word_count: 812
asset_id: microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74
moniker_range_name: 
monikers: []
item_type: Content
source_path: documentation/microsoft_general_purpose_programming_languages/MS-VBAL/d5418146-0bd2-45eb-9c7a-fd9502722c74.md
cmProducts:
- https://microsoft-devrel.poolparty.biz/DevRelOfferingOntology/c6f99e62-1cf6-4b71-af9b-649b05f80cce
spProducts:
- https://microsoft-devrel.poolparty.biz/DevRelOfferingOntology/3f56b378-07a9-4fa1-afe8-9889fdc77628
platformId: 3a63f126-d5ef-80c3-738a-6e99debdcce9
---

# [MS-VBAL]: VBA Language Specification | Microsoft Learn

Specifies the VBA Language, which defines the implementation-independent and operating system-independent programming language that is required to be supported by all conforming VBA implementations. This specification also defines all features and behaviors of the language that are required to exist and behave identically in all conforming implementations.

This page and associated content may be updated frequently. We recommend you subscribe to the [RSS feed](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d.rss) to receive update notifications.

## Published Version

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 5/20/2025 | 2.4 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-250520.docx) |

[Click here to download a zip file of all PDF files for Microsoft General Purpose Programming Languages.](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/Zip_Files/GeneralPurposeProgrammingLanguages.zip)

## Previous Versions

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 2/18/2025 | 2.3 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-250218.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-250218.docx) |
| 11/12/2024 | 2.2 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-241112.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-241112.docx) |
| 5/21/2024 | 2.1 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240521.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240521.docx) |
| 4/16/2024 | 2.0 | Major | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240416.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240416.docx) |
| 2/20/2024 | 1.8 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240220.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-240220.docx) |
| 2/16/2021 | 1.7 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-210216.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-210216.docx) |
| 11/17/2020 | 1.6 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-201117.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-201117.docx) |
| 8/18/2020 | 1.5 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-200818.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-200818.docx) |
| 9/24/2019 | 1.4 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-190924.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-190924.docx) |
| 6/18/2019 | 1.3 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-190618.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-190618.docx) |
| 12/15/2016 | 1.02 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-161215.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-161215.docx) |
| 4/30/2014 | 1.02 | Editorial | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-140430.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-VBAL/%5bMS-VBAL%5d-140430.doc) |
| 3/15/2012 | 1.01 | Major |  |
| 3/15/2010 | 1.0 | Major |  |
| 6/30/2009 | 0.95 | Major |  |
| 6/30/2008 | 0.9 | Major |  |

## Preview Versions

From time to time, Microsoft may publish a preview, or pre-release, version of an Open Specifications technical document for community review and feedback. To submit feedback for a preview version of a technical document, please follow any instructions specified for that document. If no instructions are indicated for the document, please provide feedback by using the [Open Specification Forums](https://aka.ms/AA9oo6c).

The preview period for a technical document varies. Additionally, not every technical document will be published for preview.

A preview version of this document may be available on the [Microsoft General Purpose Programming Languages](../ms-proglanglp/a0fa63f3-19ba-4525-a22f-f0ab65dc015d#previewspecs) page. After the preview period, the most current version of the document is available on this page.

## Development Resources

Find resources for creating interoperable solutions for Microsoft software, services, hardware, and non-Microsoft products: 

[Plugfests and Events](https://msdn.microsoft.com/en-us/openspecifications/dn750988), [Test Tools](https://msdn.microsoft.com/en-us/openspecifications/dn750986), [Development Support](https://msdn.microsoft.com/en-us/openspecifications/cc816063), and [Open Specifications Dev Center](https://msdn.microsoft.com/en-us/openspecifications).

## Intellectual Property Rights Notice for Open Specifications Documentation

- **Technical Documentation. **Microsoft publishes Open Specifications documentation (“this documentation”) for protocols, file formats, data portability, computer languages, and standards support. Additionally, overview documents cover inter-protocol relationships and interactions.
- **Copyrights**. This documentation is covered by Microsoft copyrights. Regardless of any other terms that are contained in the terms of use for the Microsoft website that hosts this documentation, you can make copies of it in order to develop implementations of the technologies that are described in this documentation and can distribute portions of it in your implementations that use these technologies or in your documentation as necessary to properly document the implementation. You can also distribute in your implementation, with or without modification, any schemas, IDLs, or code samples that are included in the documentation. This permission also applies to any documents that are referenced in the Open Specifications documentation.
- **No Trade Secrets**. Microsoft does not claim any trade secret rights in this documentation.
- **Patents**. Microsoft has patents that might cover your implementations of the technologies described in the Open Specifications documentation. Neither this notice nor Microsoft's delivery of this documentation grants any licenses under those patents or any other Microsoft patents. However, a given Open Specifications document might be covered by the Microsoft [Open Specifications Promise](https://go.microsoft.com/fwlink/?LinkId=214445) or the [Microsoft Community Promise](https://go.microsoft.com/fwlink/?LinkId=214448). If you would prefer a written license, or if the technologies described in this documentation are not covered by the Open Specifications Promise or Community Promise, as applicable, patent licenses are available by contacting iplg@microsoft.com.
- **License Programs**. To see all of the protocols in scope under a specific license program and the associated patents, visit the [Patent Map](https://aka.ms/AA9ufj8).
- **Trademarks**. The names of companies and products contained in this documentation might be covered by trademarks or similar intellectual property rights. This notice does not grant any licenses under those rights. For a list of Microsoft trademarks, visit https://www.microsoft.com/trademarks.
- **Fictitious Names**. The example companies, organizations, products, domain names, email addresses, logos, people, places, and events that are depicted in this documentation are fictitious. No association with any real company, organization, product, domain name, email address, logo, person, place, or event is intended or should be inferred.

**Reservation of Rights**. All other rights are reserved, and this notice does not grant any rights other than as specifically described above, whether by implication, estoppel, or otherwise. 

**Tools**. The Open Specifications documentation does not require the use of Microsoft programming tools or programming environments in order for you to develop an implementation. If you have access to Microsoft programming tools and environments, you are free to take advantage of them. Certain Open Specifications documents are intended for use in conjunction with publicly available standards specifications and network programming art and, as such, assume that the reader either is familiar with the aforementioned material or has immediate access to it.

**Support.** For questions and support, please contact dochelp@microsoft.com.