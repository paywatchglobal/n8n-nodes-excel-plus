import esbuild from 'esbuild';
import { cp, mkdir, readFile, readdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = dirname(fileURLToPath(new URL('..', import.meta.url + '/')));
const watch = process.argv.includes('--watch');

const pkg = JSON.parse(await readFile(join(projectRoot, 'package.json'), 'utf8'));
const external = [
	...Object.keys(pkg.peerDependencies ?? {}),
	...Object.keys(pkg.dependencies ?? {}),
];

const nodeEntries = (pkg.n8n?.nodes ?? []).map((distPath) => {
	const sourcePath = distPath
		.replace(/^dist\//, '')
		.replace(/\.node\.js$/, '.node.ts');
	return { src: sourcePath, out: distPath };
});

const credentialEntries = (pkg.n8n?.credentials ?? []).map((distPath) => {
	const sourcePath = distPath
		.replace(/^dist\//, '')
		.replace(/\.credentials\.js$/, '.credentials.ts');
	return { src: sourcePath, out: distPath };
});

const buildOptions = {
	bundle: true,
	platform: 'node',
	target: 'node18',
	format: 'cjs',
	external,
	sourcemap: true,
	logLevel: 'info',
};

async function copyStaticFiles(srcDir, destDir) {
	let entries;
	try {
		entries = await readdir(srcDir, { withFileTypes: true });
	} catch (error) {
		if (error.code === 'ENOENT') return;
		throw error;
	}
	await mkdir(destDir, { recursive: true });
	for (const entry of entries) {
		const src = join(srcDir, entry.name);
		const dest = join(destDir, entry.name);
		if (entry.isDirectory()) {
			await copyStaticFiles(src, dest);
			continue;
		}
		if (/\.(svg|png|json)$/i.test(entry.name)) {
			await cp(src, dest);
		}
	}
}

async function copyAllStatic() {
	for (const entry of [...nodeEntries, ...credentialEntries]) {
		const srcDir = dirname(entry.src);
		const destDir = dirname(entry.out);
		await copyStaticFiles(srcDir, destDir);
	}
}

async function buildOne(entry) {
	const options = {
		...buildOptions,
		entryPoints: [entry.src],
		outfile: entry.out,
	};
	if (watch) {
		const ctx = await esbuild.context(options);
		await ctx.watch();
		return ctx;
	}
	await esbuild.build(options);
	return null;
}

const entries = [...nodeEntries, ...credentialEntries];
if (entries.length === 0) {
	console.warn('No node or credential entries found in package.json#n8n.');
}

await Promise.all(entries.map((entry) => buildOne(entry)));
await copyAllStatic();

if (watch) {
	console.log('Watching for changes...');
} else {
	console.log(`Built ${entries.length} bundle(s).`);
}
