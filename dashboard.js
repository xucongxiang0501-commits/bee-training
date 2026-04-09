(() => {
  const scriptParts = [
    "./dashboard.runtime.part1.js",
    "./dashboard.runtime.part2.js",
    "./dashboard.runtime.part3.js",
    "./dashboard.runtime.part4.js"
  ];

  Promise.all(scriptParts.map(async (path) => {
    const response = await fetch(path, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`脚本分片加载失败: ${path}`);
    }
    return response.text();
  })).then((sources) => {
    const script = document.createElement("script");
    script.text = sources.join("\n");
    document.body.appendChild(script);
  }).catch((error) => {
    console.error(error);
    const toast = document.getElementById("toast");
    if (toast) {
      toast.textContent = error.message || "脚本加载失败";
      toast.classList.remove("translate-y-4", "opacity-0", "bg-ink");
      toast.classList.add("translate-y-0", "opacity-100", "bg-rose-700");
    }
  });
})();
