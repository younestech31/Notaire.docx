import { register } from 'node:module';
import { pathToFileURL } from 'node:url';
import fs from 'node:fs';
import path from 'node:path';

export async function resolve(specifier, context, nextResolve) {
  if (specifier.startsWith('@/')) {
    const mapped = path.resolve(process.cwd(), specifier.slice(2));
    if (fs.existsSync(`${mapped}.ts`)) {
      return nextResolve(pathToFileURL(`${mapped}.ts`).href, context);
    }
    if (fs.existsSync(`${mapped}.tsx`)) {
      return nextResolve(pathToFileURL(`${mapped}.tsx`).href, context);
    }
  }
  if (
    (specifier.startsWith('./') || specifier.startsWith('../')) &&
    !path.extname(specifier) &&
    context.parentURL
  ) {
    const parentDir = path.dirname(new URL(context.parentURL).pathname);
    const candidateTs = path.resolve(parentDir, `${specifier}.ts`);
    if (fs.existsSync(candidateTs)) {
      return nextResolve(pathToFileURL(candidateTs).href, context);
    }
    const candidateTsx = path.resolve(parentDir, `${specifier}.tsx`);
    if (fs.existsSync(candidateTsx)) {
      return nextResolve(pathToFileURL(candidateTsx).href, context);
    }
  }
  return nextResolve(specifier, context);
}

register(import.meta.url, pathToFileURL('./'));
