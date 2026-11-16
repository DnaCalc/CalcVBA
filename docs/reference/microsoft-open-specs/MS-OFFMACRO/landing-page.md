---
layout: Conceptual
title: '[MS-OFFMACRO]: Office Macro-Enabled File Format | Microsoft Learn'
canonicalUrl: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro/86fa5ba8-69cf-4164-8559-c5af66d2990e
ms.service: openspecs-office
ROBOTS: INDEX, FOLLOW
uhfHeaderId: MSDocsHeader-OpenSpecs
ms.topic: reference
ms.author: cindyle
protocol_rendering: true
description: Specifies the structures of macro-enabled WordprocessingML, SpreadsheetML, and PresentationML file formats that are either
locale: en-us
author: mrsgit09
document_id: 07801768-b2d5-2281-b16c-7cf7edd43fb6
document_version_independent_id: 31a2d968-52b6-5ede-04fc-37adb32b9320
updated_at: 2025-07-11T20:45:00.0000000Z
original_content_git_url: https://github.com/MicrosoftDocs/open_specs_office/blob/live/documentation/office_standards/MS-OFFMACRO/86fa5ba8-69cf-4164-8559-c5af66d2990e.md
gitcommit: https://github.com/MicrosoftDocs/open_specs_office/blob/f540c05614144fa39bb40491a8e3aedb3e900f78/documentation/office_standards/MS-OFFMACRO/86fa5ba8-69cf-4164-8559-c5af66d2990e.md
git_commit_id: f540c05614144fa39bb40491a8e3aedb3e900f78
site_name: Docs
depot_name: MSDN.open_specs_office
page_type: conceptual
toc_rel: toc.json
feedback_system: None
feedback_product_url: ''
feedback_help_link_type: ''
feedback_help_link_url: ''
word_count: 818
asset_id: office_standards/ms-offmacro/86fa5ba8-69cf-4164-8559-c5af66d2990e
moniker_range_name: 
monikers: []
item_type: Content
source_path: documentation/office_standards/MS-OFFMACRO/86fa5ba8-69cf-4164-8559-c5af66d2990e.md
cmProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/60932d05-feee-4685-a73b-595e25dd9318
- https://microsoft-devrel.poolparty.biz/DevRelOfferingOntology/c6f99e62-1cf6-4b71-af9b-649b05f80cce
spProducts:
- https://authoring-docs-microsoft.poolparty.biz/devrel/c43bb58b-1190-419b-8d18-e6052371b599
- https://microsoft-devrel.poolparty.biz/DevRelOfferingOntology/3f56b378-07a9-4fa1-afe8-9889fdc77628
platformId: e72d9dfd-e0af-5055-61bf-1c2c3c23c0d7
---

# [MS-OFFMACRO]: Office Macro-Enabled File Format | Microsoft Learn

Specifies the structures of macro-enabled WordprocessingML, SpreadsheetML, and PresentationML file formats that are either not present in or modified from non-macro-enabled WordprocessingML, SpreadsheetML, and PresentationML file formats.

This page and associated content may be updated frequently. We recommend you subscribe to the [RSS feed](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d.rss) to receive update notifications.

## Published Version

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 9/14/2016 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-160914.docx) |

[Click here to download a zip file of all PDF files for Word, Excel, and PowerPoint Standards Support.](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/Zip_Files/MSOFFSTAND.zip)

## Previous Versions

| Date | Protocol Revision | Revision Class | Downloads |
| --- | --- | --- | --- |
| 6/23/2016 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-160623.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-160623.docx) |
| 10/30/2014 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-141030.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-141030.doc) |
| 7/31/2014 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140731.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140731.doc) |
| 4/30/2014 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140430.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140430.doc) |
| 2/10/2014 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140210.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-140210.doc) |
| 11/18/2013 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-131118.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-131118.doc) |
| 7/30/2013 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-130730.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-130730.doc) |
| 2/11/2013 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-130211.pdf) | [DOCX](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-130211.doc) |
| 10/8/2012 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-121008.pdf) |
| 7/16/2012 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-120716.pdf) |
| 4/11/2012 | 1.5 | None | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-120411.pdf) |
| 1/20/2012 | 1.5 | Minor | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO/%5bMS-OFFMACRO%5d-120120.pdf) | [PDF](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OFFMACRO2/%5bMS-OFFMACRO2%5d-120120.pdf) |
| 6/10/2011 | 1.04 | None |  |
| 3/18/2011 | 1.04 | None |  |
| 12/17/2010 | 1.04 | None |  |
| 11/15/2010 | 1.04 | None |  |
| 9/27/2010 | 1.04 | None |  |
| 7/23/2010 | 1.04 | None |  |
| 6/29/2010 | 1.04 | Editorial |  |
| 6/7/2010 | 1.03 | Editorial |  |
| 4/30/2010 | 1.02 | Editorial |  |
| 3/31/2010 | 1.01 | Editorial |  |
| 2/19/2010 | 1.0 | Editorial |  |
| 11/6/2009 | 0.3 | Editorial |  |
| 8/28/2009 | 0.2 | Editorial |  |
| 7/13/2009 | 0.1 | Major |  |

## Preview Versions

From time to time, Microsoft may publish a preview, or pre-release, version of an Open Specifications technical document for community review and feedback. To submit feedback for a preview version of a technical document, please follow any instructions specified for that document. If no instructions are indicated for the document, please provide feedback by using the [Open Specification Forums](https://aka.ms/AA9oo6c).

The preview period for a technical document varies. Additionally, not every technical document will be published for preview.

A preview version of this document may be available on the [Word, Excel, and PowerPoint Standards Support](../ms-offstandlp/d5784a8b-7070-466b-befa-b7bf3724c6f0#previewspecs) page. After the preview period, the most current version of the document is available on this page.

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