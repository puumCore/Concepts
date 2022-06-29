package edu.puumCore.concepts._interfaces;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @author Puum Core (Mandela Muriithi)<br>
 * <a href = "https://github.com/puumCore">GitHub: Mandela Muriithi</a>
 * @version 1.0
 * @since 29/06/2022
 */

public interface Excel {

    void write_to_file(File targetFile) throws IOException;

    void read_file(File targetFile) throws IOException;

    File create_table(File targetFile) throws IOException;

}
