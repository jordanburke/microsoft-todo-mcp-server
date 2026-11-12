import { realpathSync } from "fs"
import { fileURLToPath } from "url"

/**
 * Whether the module at `moduleUrl` is the script node was started with. Compares real
 * paths, because npm runs a bin through a symlink (so argv[1] is the link, not the file),
 * and a URL percent-encodes spaces and uses forward slashes on Windows.
 */
export const isMain = (moduleUrl: string, entry: string | undefined = process.argv[1]): boolean => {
  if (!entry) return false
  try {
    return realpathSync(entry) === realpathSync(fileURLToPath(moduleUrl))
  } catch {
    return false
  }
}
