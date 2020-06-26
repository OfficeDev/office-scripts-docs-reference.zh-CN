---
title: Office 脚本 API 参考
description: Office 脚本异步 JavaScript Api 概述。
ms.date: 06/17/2020
ms.openlocfilehash: b9fa59764b7cb54567adbe05b9671bae9b232972
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44883141"
---
# <a name="office-scripts-async-api-reference"></a><span data-ttu-id="057b4-103">Office 脚本异步 API 参考</span><span class="sxs-lookup"><span data-stu-id="057b4-103">Office Scripts Async API reference</span></span>

<span data-ttu-id="057b4-104">Office 脚本异步 API 支持在 Office 脚本预览阶段执行较旧的脚本。</span><span class="sxs-lookup"><span data-stu-id="057b4-104">The Office Scripts Async API supports older scripts made during the Office Scripts preview phase.</span></span> <span data-ttu-id="057b4-105">使用此参考文档可详细了解这些旧脚本使用的类、方法和其他类型。</span><span class="sxs-lookup"><span data-stu-id="057b4-105">Use this reference documentation to learn more about the classes, methods, and other types used by these older scripts.</span></span> <span data-ttu-id="057b4-106">可通过 Office 脚本访问的所有对象都可以在页面左侧的目录中找到。</span><span class="sxs-lookup"><span data-stu-id="057b4-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="057b4-107">强烈建议使用标准 Office 脚本 Api 创建新的脚本。</span><span class="sxs-lookup"><span data-stu-id="057b4-107">We strongly recommend creating new scripts with the standard Office Scripts APIs.</span></span> <span data-ttu-id="057b4-108">如果要进行或编辑新脚本，请切换到[非异步版本](?view=office-scripts)的 api。</span><span class="sxs-lookup"><span data-stu-id="057b4-108">If you're making or editing a new script, please switch to the [non-async version](?view=office-scripts) of the APIs.</span></span>

## <a name="common-classes"></a><span data-ttu-id="057b4-109">公共类</span><span class="sxs-lookup"><span data-stu-id="057b4-109">Common classes</span></span>

<span data-ttu-id="057b4-110">以下列表细分了 Office 脚本对象模型的基础知识。</span><span class="sxs-lookup"><span data-stu-id="057b4-110">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="057b4-111">这将显示常见类以及它们之间的关系。</span><span class="sxs-lookup"><span data-stu-id="057b4-111">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="057b4-112">[工作簿](/javascript/api/office-scripts/excel/excelscript.workbook)在[WorksheetCollection](/javascript/api/office-scripts/excel/excelscript.worksheetcollection)中包含一个或多个[工作表](/javascript/api/office-scripts/excel/excelscript.worksheet)。</span><span class="sxs-lookup"><span data-stu-id="057b4-112">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet) in a [WorksheetCollection](/javascript/api/office-scripts/excel/excelscript.worksheetcollection).</span></span>
- <span data-ttu-id="057b4-113">[Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) 可通过 [Range](/javascript/api/office-scripts/excel/excelscript.range) 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="057b4-113">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="057b4-114">[Range](/javascript/api/office-scripts/excel/excelscript.range) 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="057b4-114">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="057b4-115">[Range](/javascript/api/office-scripts/excel/excelscript.range) 用于创建和放置 [Table](/javascript/api/office-scripts/excel/excelscript.table)、[Chart](/javascript/api/office-scripts/excel/excelscript.chart) 和 [Shape](/javascript/api/office-scripts/excel/excelscript.shape) 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="057b4-115">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="057b4-116">[工作表](/javascript/api/office-scripts/excel/excelscript.worksheet)包含单个工作表中存在的这些数据对象（如[ChartCollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)）的集合。</span><span class="sxs-lookup"><span data-stu-id="057b4-116">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains collections of those data objects (such as a [ChartCollection](/javascript/api/office-scripts/excel/excelscript.chartcollection)) that are present in the individual sheet.</span></span>
- <span data-ttu-id="057b4-117">[工作簿](/javascript/api/office-scripts/excel/excelscript.workbook)包含整个[工作簿](/javascript/api/office-scripts/excel/excelscript.workbook)的一些数据对象（如[TableCollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)）的集合。</span><span class="sxs-lookup"><span data-stu-id="057b4-117">[Workbooks](/javascript/api/office-scripts/excel/excelscript.workbook) contain collections of some of those data objects (such as a [TableCollection](/javascript/api/office-scripts/excel/excelscript.tablecollection)) for the entire [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook).</span></span>

<span data-ttu-id="057b4-118">有关 Office 脚本对象模型的详细信息，请访问[web 上 Excel 中 Office 脚本的脚本基础](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="057b4-118">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="057b4-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="057b4-119">See also</span></span>

- [<span data-ttu-id="057b4-120">使用 Office 脚本异步 Api 支持旧版脚本</span><span class="sxs-lookup"><span data-stu-id="057b4-120">Using the Office Scripts Async APIs to support legacy scripts</span></span>](/office/dev/scripts/develop/excel-async-model)
- [<span data-ttu-id="057b4-121">关于 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="057b4-121">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="057b4-122">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="057b4-122">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="057b4-123">Excel 网页版中 Office 脚本的脚本基础</span><span class="sxs-lookup"><span data-stu-id="057b4-123">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
