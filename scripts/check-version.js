const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "20" && nodeVersion !== "18") {
  console.error("Node version must be 20 or 18");
  process.exitCode = 1;
}
