export async function* readFileLines(file: File): AsyncGenerator<string> {
  const lineChunkGenerator = readFileLineChunks(file);
  for await (const lines of lineChunkGenerator) {
    yield* lines;
  }
}

/**
 * An async generator function that reads a file and yields batches of
 * complete lines.
 *
 * @param file The file to read.
 * @returns An AsyncGenerator that yields arrays of strings (lines).
 */
async function* readFileLineChunks(file: File): AsyncGenerator<string[]> {
  const decoder = new TextDecoder("latin1");
  const reader = file.stream().getReader();
  let remainder = "";

  try {
    while (true) {
      const { value, done } = await reader.read();

      if (done) {
        // The stream has ended. If there's any leftover text, yield it as the final batch.
        if (remainder) {
          yield [remainder];
        }
        // Exit the loop and end the generator.
        break;
      }

      const chunkText = remainder + decoder.decode(value);
      const lastNewlineIndex = chunkText.lastIndexOf("\n");

      if (lastNewlineIndex !== -1) {
        const completeLinesText = chunkText.substring(0, lastNewlineIndex);
        remainder = chunkText.substring(lastNewlineIndex + 1);

        // "Pause" and yield the batch of complete lines.
        yield completeLinesText.split("\n");
      } else {
        // No newline found, so the entire chunk is a remainder.
        // The loop will continue to get the next chunk.
        remainder = chunkText;
      }
    }
  } finally {
    // It's good practice to release the lock on the stream when you're done.
    reader.releaseLock();
  }
}
