# <a name="office-scripts-api-documentation-tools"></a><span data-ttu-id="f8aa4-101">Office脚本 API 文档工具</span><span class="sxs-lookup"><span data-stu-id="f8aa4-101">Office Scripts API Documentation Tools</span></span>

<span data-ttu-id="f8aa4-102">这些工具可帮助支持Office文档及其后面的团队。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-102">These tools help support the Office SCripts documentation and the team behind it.</span></span> <span data-ttu-id="f8aa4-103">按照以下说明运行此文件夹中的工具。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-103">Follow these instructions to run the tools in this folder.</span></span>

## <a name="coverage-tester"></a><span data-ttu-id="f8aa4-104">coverage-tester</span><span class="sxs-lookup"><span data-stu-id="f8aa4-104">coverage-tester</span></span>

<span data-ttu-id="f8aa4-105">此工具概述了每个 API 的文档覆盖范围。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-105">This tool gives an overview of the documentation coverage for each API.</span></span> <span data-ttu-id="f8aa4-106">评估每个 API 的文档质量和示例代码是否存在。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-106">Each API is assessed for documentation quality and the presence of sample code.</span></span> <span data-ttu-id="f8aa4-107">质量指标仍在开发中。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-107">The quality metrics are still in development.</span></span>

<span data-ttu-id="f8aa4-108">此工具的输出是 `.csv` 一个文件。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-108">The output of this tool is a `.csv` file.</span></span>

### <a name="coverage-tester-instructions"></a><span data-ttu-id="f8aa4-109">coverage-tester 说明</span><span class="sxs-lookup"><span data-stu-id="f8aa4-109">coverage-tester Instructions</span></span>

1. <span data-ttu-id="f8aa4-110">克隆或分叉存储库。</span><span class="sxs-lookup"><span data-stu-id="f8aa4-110">Clone or fork the repo.</span></span>
1. <span data-ttu-id="f8aa4-111">在命令窗口中，转到 `/office-scripts-docs-reference/generate-docs/tools`</span><span class="sxs-lookup"><span data-stu-id="f8aa4-111">In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`</span></span>
1. <span data-ttu-id="f8aa4-112">运行 `npm install`</span><span class="sxs-lookup"><span data-stu-id="f8aa4-112">Run `npm install`</span></span>
1. <span data-ttu-id="f8aa4-113">运行 `npm run build`</span><span class="sxs-lookup"><span data-stu-id="f8aa4-113">Run `npm run build`</span></span>
1. <span data-ttu-id="f8aa4-114">运行 `node coverage-tester`</span><span class="sxs-lookup"><span data-stu-id="f8aa4-114">Run `node coverage-tester`</span></span>
1. <span data-ttu-id="f8aa4-115">打开"API 覆盖范围Report.csv"</span><span class="sxs-lookup"><span data-stu-id="f8aa4-115">Open “API Coverage Report.csv”</span></span>
