Attribute VB_Name = "modImages_EXIF_WIA"
'@Lang VBA

  																					                                                                                            ' _
   |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                          ' _
   ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                          ' _
   ||||||||||||||||||||||||||      EXIF PROPERTIES  (WIA) V1.1      ||||||||||||||||||||||||||||||||||                                                                          ' _
   ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                          ' _
   |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                          ' _
															         					                                                                                        ' _
   AUTHOR:   Dan_W and Kallun Willock                                                                     							                                            ' _
   PURPOSE:  Routines to read and write EXIF properties to JPG image files using                                                                                                ' _
             the Windows Image Acquisition (WIA) COM Object.									 					                                                            ' _
   LICENSE:  MIT
																 	         	         	                                                                                    ' _
   VERSION:  1.1   09/04/2022      Added 400+ additional EXIF MetaData Property Tags; New subroutine                                                                            ' _
                                   to load table of EXIF MetaData Properites (EXIFTOOL.ORG) for easy reference                   	         	         	                    ' _
             1.0   25/03/2022      Completed code for OP -  									 	         	         	                                                    ' _
                                   https://www.mrexcel.com/board/threads/using-a-userform-to-change-the-document-properties-or-tags.1198206/
																                                                                                                                ' _
   USAGE:             									                                                                                                                        ' _
             WriteEXIFData Filename, PropertyName, PropertyValue, (Opt) WriteOverOriginal = True, (Opt) CreateBackup             					                            ' _
             - WriteEXIFData "C:\Temp\IMG01 - Copy.jpg", EXIFImageTitle, "New Image Title"                                                                                      ' _
             - WriteEXIFData "C:\Temp\IMG01 - Copy.jpg", EXIFImageTitle, "Live Life On The Edge", True                                                                          ' _
															         					                                                                                        ' _
             - Comments = GetEXIFData("C:\Temp\IMG_20210508_170154.jpg", EXIFImageComments)
                                                         
   Option Explicit
												                                                                     	                                                        ' _
       Property Descriptions sourced from                                                                                                                                     	' _
       https://docs.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-item-descriptions
   
   Public Enum PropertyNameEnum
       EXIFImageDateTimeOriginal = 36867
       EXIFImageTitle = 40091
       EXIFImageComments = 40092
       EXIFImageAuthor = 40093
       EXIFImageKeywords = 40094
       EXIFImageSubject = 40095
       GPSVer = 0                                        ' Version of the Global Positioning Systems (GPS) IFD, given as 2.0.0.0. This tag is mandatory when the GpsIFD tag is present. When the version is 2.0.0.0, the tag value is 0x02000000.
       GPSLatitudeRef = 1                                ' Null-terminated character string that specifies whether the latitude is north or south.: N: specifies north latitude, and: S: specifies south latitude.
       GPSLatitude = 2                                   ' Latitude. Latitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is dd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is dd/1, mmmm/100, 0/1.
       GPSLongitudeRef = 3                               ' Null-terminated character string that specifies whether the longitude is east or west longitude.: E: specifies east longitude, and: W: specifies west longitude.
       GPSLongitude = 4                                  ' Longitude. Longitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes and seconds are expressed, the format is ddd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is ddd/1, mmmm/100, 0/1.
       GPSAltitudeRef = 5                                ' Reference altitude, in meters.
       GPSAltitude = 6                                   ' Altitude, in meters, based on the reference altitude specified by GpsAltitudeRef.
       GPSGPSTime = 7                                    ' Time as Coordinated Universal Time (UTC). The value is expressed as three rational numbers that give the hour, minute, and second.
       GPSGPSSatellites = 8                              ' Null-terminated character string that specifies the GPS satellites used for measurements. This tag can be used to specify the ID number, angle of elevation, azimuth, SNR, and other information about each satellite. The format is not specified. If the GPS receiver is incapable of taking measurements, the value of the tag must be set to: NULL.
       GPSGPSStatus = 9                                  ' Null-terminated character string that specifies the status of the GPS receiver when the image is recorded.: A: means measurement is in progress, and: V: means the measurement is Interoperability.
       GPSGPSMeasureMode = 10                            ' Null-terminated character string that specifies the GPS measurement mode.: 2: specifies 2-D measurement, and: 3: specifies 3-D measurement.
       GPSGPSDop = 11                                    ' GPS DOP (data degree of precision). An HDOP value is written during 2-D measurement, and a PDOP value is written during 3-D measurement.
       GPSSpeedRef = 12                                  ' Null-terminated character string that specifies the unit used to express the GPS receiver speed of movement.: K,: M, and: N: represent kilometers per hour, miles per hour, and knots respectively.
       GPSSpeed = 13                                     ' Speed of the GPS receiver movement.
       GPSTrackRef = 14                                  ' Null-terminated character string that specifies the reference for giving the direction of GPS receiver movement.: T: specifies true direction, and: M: specifies magnetic direction.
       GPSTrack = 15                                     ' Direction of GPS receiver movement. The range of values is from 0.00 to 359.99.
       GPSImgDirRef = 16                                 ' Null-terminated character string that specifies the reference for the direction of the image when it is captured.: T: specifies true direction, and: M: specifies magnetic direction.
       GPSImgDir = 17                                    ' Direction of the image when it was captured. The range of values is from 0.00 to 359.99.
       GPSMapDatum = 18                                  ' Null-terminated character string that specifies geodetic survey data used by the GPS receiver. If the survey data is restricted to Japan, the value of this tag is: TOKYO: or: WGS-84.
       GPSDestLatRef = 19                                ' Null-terminated character string that specifies whether the latitude of the destination point is north or south latitude.: N: specifies north latitude, and: S: specifies south latitude.
       GPSDestLat = 20                                   ' Latitude of the destination point. The latitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is dd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is dd/1, mmmm/100, 0/1.
       GPSDestLongRef = 21                               ' Null-terminated character string that specifies whether the longitude of the destination point is east or west longitude.: E: specifies east longitude, and: W: specifies west longitude.
       GPSDestLong = 22                                  ' Longitude of the destination point. The longitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is ddd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is ddd/1, mmmm/100, 0/1.
       GPSDestBearRef = 23                               ' Null-terminated character string that specifies the reference used for giving the bearing to the destination point.: T: specifies true direction, and: M: specifies magnetic direction.
       GPSDestBear = 24                                  ' Bearing to the destination point. The range of values is from 0.00 to 359.99.
       GPSDestDistRef = 25                               ' Null-terminated character string that specifies the unit used to express the distance to the destination point. K, M, and N represent kilometers, miles, and knots respectively.
       GPSDestDist = 26                                  ' Distance to the destination point.
       NewSubfileType = 254                              ' Type of data in a subfile.
       SubfileType = 255                                 ' Type of data in a subfile.
       ImageWidth = 256                                  ' Number of pixels per row.
       ImageHeight = 257                                 ' Number of pixel rows.
       BitsPerSample = 258                               ' Number of bits per color component. See also: SamplesPerPixel.
       Compression = 259                                 ' Compression scheme used for the image data.
       PhotometricInterp = 262                           ' How pixel data will be interpreted.
       ThreshHolding = 263                               ' Technique used to convert from gray pixels to black and white pixels.
       CellWidth = 264                                   ' Width of the dithering or halftoning matrix.
       CellHeight = 265                                  ' Height of the dithering or halftoning matrix.
       FillOrder = 266                                   ' Logical order of bits in a byte.
       DocumentName = 269                                ' Null-terminated character string that specifies the name of the document from which the image was scanned.
       ImageDescription = 270                            ' Null-terminated character string that specifies the title of the image.
       EquipMake = 271                                   ' Null-terminated character string that specifies the manufacturer of the equipment used to record the image.
       EquipModel = 272                                  ' Null-terminated character string that specifies the model name or model number of the equipment used to record the image.
       StripOffsets = 273                                ' For each strip, the byte offset of that strip. See also: RowsPerStrip: and: StripBytesCount.
       Orientation = 274                                 ' Image orientation viewed in terms of rows and columns.
       SamplesPerPixel = 277                             ' Number of color components per pixel.
       RowsPerStrip = 278                                ' Number of rows per strip. See also: StripBytesCount: and: StripOffsets.
       StripBytesCount = 279                             ' For each strip, the total number of bytes in that strip.
       MinSampleValue = 280                              ' For each color component, the minimum value assigned to that component. See also: SamplesPerPixel.
       MaxSampleValue = 281                              ' For each color component, the maximum value assigned to that component. See also: SamplesPerPixel.
       XResolution = 282                                 ' Number of pixels per unit in the image width (x) direction. The unit is specified by: ResolutionUnit.
       YResolution = 283                                 ' Number of pixels per unit in the image height (y) direction. The unit is specified by: ResolutionUnit.
       PlanarConfig = 284                                ' Whether pixel components are recorded in chunky or planar format.
       PageName = 285                                    ' Null-terminated character string that specifies the name of the page from which the image was scanned.
       XPosition = 286                                   ' Offset from the left side of the page to the left side of the image. The unit of measure is specified by: ResolutionUnit.
       YPosition = 287                                   ' Offset from the top of the page to the top of the image. The unit of measure is specified by: ResolutionUnit.
       FreeOffset = 288                                  ' For each string of contiguous unused bytes, the byte offset of that string.
       FreeByteCounts = 289                              ' For each string of contiguous unused bytes, the number of bytes in that string.
       GrayResponseUnit = 290                            ' Precision of the number specified by GrayResponseCurve. 1 specifies tenths, 2 specifies hundredths, 3 specifies thousandths, and so on.
       GrayResponseCurve = 291                           ' For each possible pixel value in a grayscale image, the optical density of that pixel value.
       T4Option = 292                                    ' Set of flags that relate to T4 encoding.
       T6Option = 293                                    ' Set of flags that relate to T6 encoding.
       ResolutionUnit = 296                              ' Unit of measure for the horizontal resolution and the vertical resolution.
       PageNumber = 297                                  ' Page number of the page from which the image was scanned.
       TransferFunction = 301                            ' Tables that specify transfer functions for the image.
       SoftwareUsed = 305                                ' Null-terminated character string that specifies the name and version of the software or firmware of the device used to generate the image.
       DateTime = 306                                    ' Date and time the image was created.
       Artist = 315                                      ' Null-terminated character string that specifies the name of the person who created the image.
       HostComputer = 316                                ' Null-terminated character string that specifies the computer and/or operating system used to create the image.
       Predictor = 317                                   ' Type of prediction scheme that was applied to the image data before the encoding scheme was applied.
       WhitePoint = 318                                  ' Chromaticity of the white point of the image.
       PrimaryChromaticities = 319                       ' For each of the three primary colors in the image, the chromaticity of that color.
       ColorMap = 320                                    ' Color palette (lookup table) for a palette-indexed image.
       HalftoneHints = 321                               ' Information used by the halftone function
       TileWidth = 322                                   ' Number of pixel columns in each tile.
       TileLength = 323                                  ' Number of pixel rows in each tile.
       TileOffset = 324                                  ' For each tile, the byte offset of that tile.
       TileByteCounts = 325                              ' For each tile, the number of bytes in that tile.
       InkSet = 332                                      ' Set of inks used in a separated image.
       InkNames = 333                                    ' Sequence of concatenated, null-terminated, character strings that specify the names of the inks used in a separated image.
       NumberOfInks = 334                                ' Number of inks.
       DotRange = 336                                    ' Color component values that correspond to a 0 percent dot and a 100 percent dot.
       TargetPrinter = 337                               ' Null-terminated character string that describes the intended printing environment.
       ExtraSamples = 338                                ' Number of extra color components. For example, one extra component might hold an alpha value.
       SampleFormat = 339                                ' For each color component, the numerical format (unsigned, signed, floating point) of that component. See also: SamplesPerPixel.
       SMinSampleValue = 342                             ' For each color component, the minimum value of that component. See also: SamplesPerPixel.
       SMaxSampleValue = 512                             ' For each color component, the maximum value of that component. See also: SamplesPerPixel.
       TransferRange = 513                               ' Table of values that extends the range of the transfer function.
       JPEGProc = 514                                    ' JPEG compression process.
       JPEGInterFormat = 515                             ' Offset to the start of a JPEG bitstream.
       JPEGInterLength = 517                             ' Length, in bytes, of the JPEG bitstream.
       JPEGRestartInterval = 518                         ' Length of the restart interval.
       JPEGLosslessPredictors = 519                      ' For each color component, a lossless predictor-selection value for that component. See also: SamplesPerPixel.
       JPEGPointTransforms = 520                         ' For each color component, a point transformation value for that component. See also: SamplesPerPixel.
       JPEGQTables = 521                                 ' For each color component, the offset to the quantization table for that component. See also: SamplesPerPixel.
       JPEGDCTables = 529                                ' For each color component, the offset to the DC Huffman table (or lossless Huffman table) for that component. See also: SamplesPerPixel.
       JPEGACTables = 530                                ' For each color component, the offset to the AC Huffman table for that component. See also: SamplesPerPixel.
       YCbCrCoefficients = 531                           ' Coefficients for transformation from RGB to YCbCr image data.
       YCbCrSubsampling = 532                            ' Sampling ratio of chrominance components in relation to the luminance component.
       YCbCrPositioning = 769                            ' Position of chrominance components in relation to the luminance component.
       REFBlackWhite = 770                               ' Reference black point value and reference white point value.
       Gamma = 771                                       ' Gamma value attached to the image. The gamma value is stored as a rational number (pair of: long) with a numerator of 100000. For example, a gamma value of 2.2 is stored as the pair (100000, 45455).
       ICCProfileDescriptor = 800                        ' Null-terminated character string that identifies an ICC profile.
       SRGBRenderingIntent = 20481                       ' Saturation intent, which is suitable for charts and graphs, preserves saturation at the expense of hue and lightness.
       ImageTitle = 20482                                ' Null-terminated character string that specifies the title of the image.
       ResolutionXUnit = 20483                           ' Units in which to display horizontal resolution.
       ResolutionYUnit = 20484                           ' Units in which to display vertical resolution.
       ResolutionXLengthUnit = 20485                     ' Units in which to display the image width.
       ResolutionYLengthUnit = 20486                     ' Units in which to display the image height.
       PrintFlags = 20487                                ' Sequence of one-byte Boolean values that specify printing options.
       PrintFlagsVersion = 20488                         ' Print flags version.
       PrintFlagsCrop = 20489                            ' Print flags center crop marks.
       PrintFlagsBleedWidth = 20490                      ' Print flags bleed width.
       PrintFlagsBleedWidthScale = 20491                 ' Print flags bleed width scale.
       HalftoneLPI = 20492                               ' Ink's screen frequency, in lines per inch.
       HalftoneLPIUnit = 20493                           ' Units for the screen frequency.
       HalftoneDegree = 20494                            ' Angle for screen.
       HalftoneShape = 20495                             ' Shape of the halftone dots.
       HalftoneMisc = 20496                              ' Miscellaneous halftone information.
       HalftoneScreen = 20497                            ' Boolean value that specifies whether to use the printer's default screens.
       JPEGQuality = 20498                               ' Private tag used by the Adobe Photoshop format. Not for public use.
       GridSize = 20499                                  ' Block of information about grids and guides.
       ThumbnailFormat = 20500                           ' Format of the thumbnail image.
       ThumbnailWidth = 20501                            ' Width, in pixels, of the thumbnail image.
       ThumbnailHeight = 20502                           ' Height, in pixels, of the thumbnail image.
       ThumbnailColorDepth = 20503                       ' bits per pixel (BPP) for the thumbnail image.
       ThumbnailPlanes = 20504                           ' Number of color planes for the thumbnail image.
       ThumbnailRawBytes = 20505                         ' Byte offset between rows of pixel data.
       ThumbnailSize = 20506                             ' Total size, in bytes, of the thumbnail image.
       ThumbnailCompressedSize = 20507                   ' Compressed size, in bytes, of the thumbnail image.
       ColorTransferFunction = 20512                     ' Table of values that specify color transfer functions.
       ThumbnailData = 20513                             ' Raw thumbnail bits in JPEG or RGB format. Depends on ThumbnailFormat.
       ThumbnailImageWidth = 20514                       ' Number of pixels per row in the thumbnail image.
       ThumbnailImageHeight = 20515                      ' Number of pixel rows in the thumbnail image.
       ThumbnailBitsPerSample = 20516                    ' Number of bits per color component in the thumbnail image. See also: ThumbnailSamplesPerPixel.
       ThumbnailCompression = 20517                      ' Compression scheme used for thumbnail image data.
       ThumbnailPhotometricInterp = 20518                ' How thumbnail pixel data will be interpreted.
       ThumbnailImageDescription = 20519                 ' Null-terminated character string that specifies the title of the image.
       ThumbnailEquipMake = 20520                        ' Null-terminated character string that specifies the manufacturer of the equipment used to record the thumbnail image.
       ThumbnailEquipModel = 20521                       ' Null-terminated character string that specifies the model name or model number of the equipment used to record the thumbnail image.
       ThumbnailStripOffsets = 20522                     ' For each strip in the thumbnail image, the byte offset of that strip. See also: ThumbnailRowsPerStrip: and: ThumbnailStripBytesCount.
       ThumbnailOrientation = 20523                      ' Thumbnail image orientation in terms of rows and columns. See also: Orientation.
       ThumbnailSamplesPerPixel = 20524                  ' Number of color components per pixel in the thumbnail image.
       ThumbnailRowsPerStrip = 20525                     ' Number of rows per strip in the thumbnail image. See also: ThumbnailStripBytesCount: and: ThumbnailStripOffsets.
       ThumbnailStripBytesCount = 20526                  ' For each thumbnail image strip, the total number of bytes in that strip.
       ThumbnailResolutionX = 20527                      ' Thumbnail resolution in the width direction. The resolution unit is given in: ThumbnailResolutionUnit.
       ThumbnailResolutionY = 20528                      ' Thumbnail resolution in the height direction. The resolution unit is given in: ThumbnailResolutionUnit.
       ThumbnailPlanarConfig = 20529                     ' Whether pixel components in the thumbnail image are recorded in chunky or planar format. See also: PlanarConfig.
       ThumbnailResolutionUnit = 20530                   ' Unit of measure for the horizontal resolution and the vertical resolution of the thumbnail image. See also: ResolutionUnit.
       ThumbnailTransferFunction = 20531                 ' Tables that specify transfer functions for the thumbnail image. See also: TransferFunction.
       ThumbnailSoftwareUsed = 20532                     ' Null-terminated character string that specifies the name and version of the software or firmware of the device used to generate the thumbnail image.
       ThumbnailDateTime = 20533                         ' Date and time the thumbnail image was created. See also: DateTime.
       ThumbnailArtist = 20534                           ' Null-terminated character string that specifies the name of the person who created the thumbnail image.
       ThumbnailWhitePoint = 20535                       ' Chromaticity of the white point of the thumbnail image. See also: WhitePoint.
       ThumbnailPrimaryChromaticities = 20536            ' For each of the three primary colors in the thumbnail image, the chromaticity of that color. See also: PrimaryChromaticities.
       ThumbnailYCbCrCoefficients = 20537                ' Coefficients for transformation from RGB to YCbCr data for the thumbnail image. See also: YCbCrCoefficients.
       ThumbnailYCbCrSubsampling = 20538                 ' Sampling ratio of chrominance components in relation to the luminance component for the thumbnail image. See also: YCbCrSubsampling.
       ThumbnailYCbCrPositioning = 20539                 ' Position of chrominance components in relation to the luminance component for the thumbnail image. See also: YCbCrPositioning.
       ThumbnailRefBlackWhite = 20624                    ' Reference black point value and reference white point value for the thumbnail image. See also: REFBlackWhite.
       ThumbnailCopyRight = 20625                        ' Null-terminated character string that contains copyright information for the thumbnail image.
       LuminanceTable = 20736                            ' Luminance table. The luminance table and the chrominance table are used to control JPEG quality. A valid luminance or chrominance table has 64 entries of type TypeShort. If an image has either a luminance table or a chrominance table, then it must have both tables.
       ChrominanceTable = 20737                          ' Chrominance table. The luminance table and the chrominance table are used to control JPEG quality. A valid luminance or chrominance table has 64 entries of type TypeShort. If an image has either a luminance table or a chrominance table, then it must have both tables.
       FrameDelay = 20738                                ' Time delay, in hundredths of a second, between two frames in an animated GIF image.
       LoopCount = 20739                                 ' For an animated GIF image, the number of times to display the animation. A value of 0 specifies that the animation should be displayed infinitely.
       GlobalPalette = 20740                             ' Color palette for an indexed bitmap in a GIF image.
       IndexBackground = 20752                           ' Index of the background color in the palette of a GIF image.
       IndexTransparent = 20753                          ' Index of the transparent color in the palette of a GIF image.
       PixelUnit = 20754                                 ' Unit for PixelPerUnitX and PixelPerUnitY.
       PixelPerUnitX = 20755                             ' Pixels per unit in the x direction.
       PixelPerUnitY = 33432                             ' Pixels per unit in the y direction.
       PaletteHistogram = 33434                          ' Palette histogram.
       Copyright = 33437                                 ' Null-terminated character string that contains copyright information.
       EXIFExposureTime = 34665                          ' Exposure time, measured in seconds.
       EXIFFNumber = 34675                               ' F number.
       EXIFIFD = 34850                                   ' Private tag used by GDI+. Not for public use. GDI+ uses this tag to locate Exif-specific information.
       ICCProfile = 34852                                ' ICC profile embedded in the image.
       EXIFExposureProg = 34853                          ' Class of the program used by the camera to set exposure when the picture is taken.
       EXIFSpectralSense = 34855                         ' Null-terminated character string that specifies the spectral sensitivity of each channel of the camera used. The string is compatible with the standard developed by the ASTM Technical Committee.
       GPSIFD = 34856                                    ' Offset to a block of GPS property items. Property items whose tags have the prefix Gps are stored in the GPS block. The GPS property items are defined in the EXIF specification. GDI+ uses this tag to locate GPS information, but GDI+ does not expose this tag for public use.
       EXIFISOSpeed = 36864                              ' ISO speed and ISO latitude of the camera or input device as specified in ISO 12232.
       EXIFOECF = 36867                                  ' Optoelectronic conversion function (OECF) specified in ISO 14524. The OECF is the relationship between the camera optical input and the image values.
       EXIFVer = 36868                                   ' Version of the EXIF standard supported. Nonexistence of this field is taken to mean nonconformance to the standard. Conformance to the standard is indicated by recording 0210 as a 4-byte ASCII string. Because the type is TypeUndefined, there is no NULL terminator.
       EXIFDTOrig = 37121                                ' Date and time when the original image data was generated. For a DSC, the date and time when the picture was taken. The format is YYYY:MM:DD HH:MM:SS with time shown in 24-hour format and the date and time separated by one blank character (0x2000). The character string length is 20 bytes including the NULL terminator. When the field is empty, it is treated as unknown.
       EXIFDTDigitized = 37122                           ' The format is YYYY:MM:DD HH:MM:SS with time shown in 24-hour format and the date and time separated by one blank character (0x2000). The character string length is 20 bytes including the NULL terminator. When the field is empty, it is treated as unknown.
       EXIFCompConfig = 37377                            ' However, because PhotometricInterp can only express the order of Y, Cb, and Cr, this tag is provided for cases when compressed data uses components other than Y, Cb, and Cr and to support other sequences.
       EXIFCompBPP = 37378                               ' Information specific to compressed data. The compression mode used for a compressed image is indicated in unit BPP.
       EXIFShutterSpeed = 37379                          ' Shutter speed. The unit is the Additive System of Photographic Exposure (APEX) value.
       EXIFAperture = 37380                              ' Lens aperture. The unit is the APEX value.
       EXIFBrightness = 37381                            ' Brightness value. The unit is the APEX value. Ordinarily it is given in the range of -99.99 to 99.99.
       EXIFExposureBias = 37382                          ' Exposure bias. The unit is the APEX value. Ordinarily it is given in the range -99.99 to 99.99.
       EXIFMaxAperture = 37383                           ' Smallest F number of the lens. The unit is the APEX value. Ordinarily it is given in the range of 00.00 to 99.99, but it is not limited to this range.
       EXIFSubjectDist = 37384                           ' Distance to the subject, measured in meters.
       EXIFMeteringMode = 37385                          ' Metering mode.
       EXIFLightSource = 37386                           ' Type of light source.
       EXIFFlash = 37500                                 ' Flash status. This tag is recorded when an image is taken using a strobe light (flash). Bit 0 indicates the flash firing status, and bits 1 and 2 indicate the flash return status.
       EXIFFocalLength = 37510                           ' Actual focal length, in millimeters, of the lens. Conversion is not made to the focal length of a 35 millimeter film camera.
       EXIFMakerNote = 37520                             ' Note tag. A tag used by manufacturers of EXIF writers to record information. The contents are up to the manufacturer.
       EXIFUserComment = 37521                           ' Comment tag. A tag used by EXIF users to write keywords or comments about the image besides those in ImageDescription and without the character-code limitations of the ImageDescription tag.
       EXIFDTSubsec = 37522                              ' Null-terminated character string that specifies a fraction of a second for the DateTime tag.
       EXIFDTOrigSS = 40960                              ' Null-terminated character string that specifies a fraction of a second for the ExifDTOrig tag.
       EXIFDTDigSS = 40961                               ' Null-terminated character string that specifies a fraction of a second for the ExifDTDigitized tag.
       EXIFFPXVer = 40962                                ' FlashPix format version supported by an FPXR file. If the FPXR function supports FlashPix format version 1.0, this is indicated similarly to ExifVer by recording 0100 as a 4-byte ASCII string. Because the type is TypeUndefined, there is no NULL terminator.
       EXIFColorSpace = 40963                            ' Color space specifier. Normally sRGB (=1) is used to define the color space based on the PC monitor conditions and environment. If a color space other than sRGB is used, Uncalibrated (=0xFFFF) is set. Image data recorded as Uncalibrated can be treated as sRGB when it is converted to FlashPix.
       EXIFPixXDim = 40964                               ' Information specific to compressed data. When a compressed file is recorded, the valid width of the meaningful image must be recorded in this tag, whether or not there is padding data or a restart marker. This tag should not exist in an uncompressed file.
       EXIFPixYDim = 40965                               ' Information specific to compressed data. When a compressed file is recorded, the valid height of the meaningful image must be recorded in this tag whether or not there is padding data or a restart marker. This tag should not exist in an uncompressed file. Because data padding is unnecessary in the vertical direction, the number of lines recorded in this valid image height tag will be the same as that recorded in the SOF.
       EXIFRelatedWav = 41483                            ' The name of an audio file related to the image data. The only relational information recorded is the EXIF audio file name and extension (an ASCII string that consists of 8 characters plus a period (.), plus 3 characters). The path is not recorded. When you use this tag, audio files must be recorded in conformance with the EXIF audio format. Writers can also store audio data within APP2 as FlashPix extension stream data.
       EXIFInterop = 41484                               ' Offset to a block of property items that contain interoperability information.
       EXIFFlashEnergy = 41486                           ' Strobe energy, in Beam Candle Power Seconds (BCPS), at the time the image was captured.
       EXIFSpatialFR = 41487                             ' Camera or input device spatial frequency table and SFR values in the image width, image height, and diagonal direction, as specified in ISO 12233.
       EXIFFocalXRes = 41488                             ' Number of pixels in the image width (x) direction per unit on the camera focal plane. The unit is specified in ExifFocalResUnit.
       EXIFFocalYRes = 41492                             ' Number of pixels in the image height (y) direction per unit on the camera focal plane. The unit is specified in ExifFocalResUnit.
       EXIFFocalResUnit = 41493                          ' Unit of measure for ExifFocalXRes and ExifFocalYRes.
       EXIFSubjectLoc = 41495                            ' Location of the main subject in the scene. The value of this tag represents the pixel at the center of the main subject relative to the left edge. The first value indicates the column number, and the second value indicates the row number.
       EXIFExposureIndex = 41728                         ' Exposure index selected on the camera or input device at the time the image was captured.
       EXIFSensingMethod = 41729                         ' Image sensor type on the camera or input device.
       EXIFFileSource = 41730                            ' The image source. If a DSC recorded the image, the value of this tag is 3.
       EXIFIFDCustomRendered = 41985
       EXIFIFDExposureMode = 41986
       EXIFIFDWhiteBalance = 41987
       EXIFIFDDigitalZoomRatio = 41988
       EXIFIFDFocalLengthIn35mmFormat = 41989
       EXIFIFDSceneCaptureType = 41990
       EXIFIFDGainControl = 41991
       EXIFIFDContrast = 41992
       EXIFIFDSaturation = 41993
       EXIFIFDSharpness = 41994
       EXIFIFDDeviceSettingDescription = 41995
       EXIFIFDSubjectDistanceRange = 41996
       EXIFIFDImageUniqueID = 42016
       EXIFIFDOwnerName = 42032
       EXIFIFDSerialNumber = 42033
       EXIFIFDLensInfo = 42034
       EXIFIFDLensMake = 42035
       EXIFIFDLensModel = 42036
       EXIFIFDLensSerialNumber = 42037
       EXIFIFDCompositeImage = 42080
       EXIFIFDCompositeImageCount = 42081
       EXIFIFDCompositeImageExposureTimes = 42082
       EXIFIFDGDALMetadata = 42112
       EXIFIFDGDALNoData = 42113
       EXIFIFDGamma = 42240
       EXIFIFDExpandSoftware = 44992
       EXIFIFDExpandLens = 44993
       EXIFIFDExpandFilm = 44994
       EXIFIFDExpandFilterLens = 44995
       EXIFIFDExpandScanner = 44996
       EXIFIFDExpandFlashLamp = 44997
       EXIFIFDHasselbladRawImage = 46275
       EXIFIFDPixelFormat = 48129
       EXIFIFDTransformation = 48130
       EXIFIFDUncompressed = 48131
       EXIFIFDImageType = 48132
       EXIFIFDImageWidth = 48256
       EXIFIFDImageHeight = 48257
       EXIFIFDWidthResolution = 48258
       EXIFIFDHeightResolution = 48259
       EXIFIFDImageOffset = 48320
       EXIFIFDImageByteCount = 48321
       EXIFIFDAlphaOffset = 48322
       EXIFIFDAlphaByteCount = 48323
       EXIFIFDImageDataDiscard = 48324
       EXIFIFDAlphaDataDiscard = 48325
       EXIFIFDOceScanjobDesc = 50215
       EXIFIFDOceApplicationSelector = 50216
       EXIFIFDOceIDNumber = 50217
       EXIFIFDOceImageLogic = 50218
       EXIFIFDAnnotations = 50255
       EXIFIFDPrintIM = 50341
       EXIFIFDHasselbladExif = 50459
       EXIFIFDOriginalFileName = 50547
       EXIFIFDUSPTOOriginalContentType = 50560
       EXIFIFDCR2CFAPattern = 50656
       EXIFIFDDNGVersion = 50706
       EXIFIFDDNGBackwardVersion = 50707
       EXIFIFDUniqueCameraModel = 50708
       EXIFIFDLocalizedCameraModel = 50709
       EXIFIFDCFAPlaneColor = 50710
       EXIFIFDCFALayout = 50711
       EXIFIFDLinearizationTable = 50712
       EXIFIFDBlackLevelRepeatDim = 50713
       EXIFIFDBlackLevel = 50714
       EXIFIFDBlackLevelDeltaH = 50715
       EXIFIFDBlackLevelDeltaV = 50716
       EXIFIFDWhiteLevel = 50717
       EXIFIFDDefaultScale = 50718
       EXIFIFDDefaultCropOrigin = 50719
       EXIFIFDDefaultCropSize = 50720
       EXIFIFDColorMatrix1 = 50721
       EXIFIFDColorMatrix2 = 50722
       EXIFIFDCameraCalibration1 = 50723
       EXIFIFDCameraCalibration2 = 50724
       EXIFIFDReductionMatrix1 = 50725
       EXIFIFDReductionMatrix2 = 50726
       EXIFIFDAnalogBalance = 50727
       EXIFIFDAsShotNeutral = 50728
       EXIFIFDAsShotWhiteXY = 50729
       EXIFIFDBaselineExposure = 50730
       EXIFIFDBaselineNoise = 50731
       EXIFIFDBaselineSharpness = 50732
       EXIFIFDBayerGreenSplit = 50733
       EXIFIFDLinearResponseLimit = 50734
       EXIFIFDCameraSerialNumber = 50735
       EXIFIFDDNGLensInfo = 50736
       EXIFIFDChromaBlurRadius = 50737
       EXIFIFDAntiAliasStrength = 50738
       EXIFIFDShadowScale = 50739
       EXIFIFDSR2Private = 50740
       EXIFIFDMakerNoteSafety = 50741
       EXIFIFDRawImageSegmentation = 50752
       EXIFIFDCalibrationIlluminant1 = 50778
       EXIFIFDCalibrationIlluminant2 = 50779
       EXIFIFDBestQualityScale = 50780
       EXIFIFDRawDataUniqueID = 50781
       EXIFIFDAliasLayerMetadata = 50784
       EXIFIFDOriginalRawFileName = 50827
       EXIFIFDOriginalRawFileData = 50828
       EXIFIFDActiveArea = 50829
       EXIFIFDMaskedAreas = 50830
       EXIFIFDAsShotICCProfile = 50831
       EXIFIFDAsShotPreProfileMatrix = 50832
       EXIFIFDCurrentICCProfile = 50833
       EXIFIFDCurrentPreProfileMatrix = 50834
       EXIFIFDColorimetricReference = 50879
       EXIFIFDSRawType = 50885
       EXIFIFDPanasonicTitle = 50898
       EXIFIFDPanasonicTitle2 = 50899
       EXIFIFDCameraCalibrationSig = 50931
       EXIFIFDProfileCalibrationSig = 50932
       EXIFIFDProfileIFD = 50933
       EXIFIFDAsShotProfileName = 50934
       EXIFIFDNoiseReductionApplied = 50935
       EXIFIFDProfileName = 50936
       EXIFIFDProfileHueSatMapDims = 50937
       EXIFIFDProfileHueSatMapData1 = 50938
       EXIFIFDProfileHueSatMapData2 = 50939
       EXIFIFDProfileToneCurve = 50940
       EXIFIFDProfileEmbedPolicy = 50941
       EXIFIFDProfileCopyright = 50942
       EXIFIFDForwardMatrix1 = 50964
       EXIFIFDForwardMatrix2 = 50965
       EXIFIFDPreviewApplicationName = 50966
       EXIFIFDPreviewApplicationVersion = 50967
       EXIFIFDPreviewSettingsName = 50968
       EXIFIFDPreviewSettingsDigest = 50969
       EXIFIFDPreviewColorSpace = 50970
       EXIFIFDPreviewDateTime = 50971
       EXIFIFDRawImageDigest = 50972
       EXIFIFDOriginalRawFileDigest = 50973
       EXIFIFDSubTileBlockSize = 50974
       EXIFIFDRowInterleaveFactor = 50975
       EXIFIFDProfileLookTableDims = 50981
       EXIFIFDProfileLookTableData = 50982
       EXIFIFDOpcodeList1 = 51008
       EXIFIFDOpcodeList2 = 51009
       EXIFIFDOpcodeList3 = 51022
       MISCNoiseProfile = 51041
       MISCTimeCodes = 51043
       MISCFrameRate = 51044
       MISCTStop = 51058
       MISCReelName = 51081
       MISCOriginalDefaultFinalSize = 51089
       MISCOriginalBestQualitySize = 51090
       MISCOriginalDefaultCropSize = 51091
       MISCCameraLabel = 51105
       MISCProfileHueSatMapEncoding = 51107
       MISCProfileLookTableEncoding = 51108
       MISCBaselineExposureOffset = 51109
       MISCDefaultBlackRender = 51110
       MISCNewRawImageDigest = 51111
       MISCRawToPreviewGain = 51112
       MISCCacheVersion = 51114
       MISCDefaultUserCrop = 51125
       MISCNikonNEFInfo = 51157
       MISCDepthFormat = 51177
       MISCDepthNear = 51178
       MISCDepthFar = 51179
       MISCDepthUnits = 51180
       MISCDepthMeasureType = 51181
       MISCEnhanceParams = 51182
       MISCProfileGainTableMap = 52525
       MISCSemanticName = 52526
       MISCSemanticInstanceIFD = 52528
       MISCCalibrationIlluminant3 = 52529
       MISCCameraCalibration3 = 52530
       MISCColorMatrix3 = 52531
       MISCForwardMatrix3 = 52532
       MISCIlluminantData1 = 52533
       MISCIlluminantData2 = 52534
       MISCIlluminantData3 = 52535
       MISCMaskSubArea = 52536
       MISCProfileHueSatMapData3 = 52537
       MISCReductionMatrix3 = 52538
       MISCRGBTables = 52539
       MISCpadding = 59932
       MISCOffsetSchema = 59933
       MISCOwnerName = 65000
       MISCSerialNumber = 65001
       MISCLens = 65002
       MISCKDC_IFD = 65024
       MISCRawFile = 65100
       MISCConverter = 65101
       MISCWhiteBalance = 65102
       MISCExposure = 65105
       MISCShadows = 65106
       MISCBrightness = 65107
       MISCContrast = 65108
       MISCSaturation = 65109
       MISCSharpness = 65110
       MISCSmoothness = 65111
       MISCMoireFilter = 65112
   End Enum
                                                         
   Private Enum WIAImagePropertyType
       UndefinedImagePropertyType = 1000
       ByteImagePropertyType = 1001
       StringImagePropertyType = 1002
       UnsignedIntegerImagePropertyType = 1003
       LongImagePropertyType = 1004
       UnsignedLongImagePropertyType = 1005
       RationalImagePropertyType = 1006
       UnsignedRationalImagePropertyType = 1007
       VectorOfUndefinedImagePropertyType = 1100
       VectorOfBytesImagePropertyType = 1101
       VectorOfUnsignedIntegersImagePropertyType = 1102
       VectorOfLongsImagePropertyType = 1103
       VectorOfUnsignedLongsImagePropertyType = 1104
       VectorOfRationalsImagePropertyType = 1105
       VectorOfUnsignedRationalsImagePropertyType = 1106
   End Enum
   
   ' Sample image: https://github.com/KallunWillock/JustMoreVBA/raw/main/Images/pexels-jill-evans-11567527.jpg
   Const TargetFileName = "C:\PATHTOFILE\pexels-jill-evans-11567527.jpg"
                                                         
   Sub Test_WriteProperties()
                                                         
       Dim NewFileName As String
       NewFileName = WriteEXIFData(TargetFileName, EXIFImageTitle, "White Concrete Building Under White Sky", True, True)
       WriteEXIFData NewFileName, EXIFImageAuthor, "Jill Evans"
       WriteEXIFData NewFileName, EXIFImageSubject, "Photo by Jill Evans from Pexels"
       WriteEXIFData NewFileName, EXIFImageComments, "Source: https://www.pexels.com/photo/white-concrete-building-under-white-sky-11567527/"
                                                         
       Debug.Print NewFileName
                                                         
   End Sub
                                                         
   Sub Test_ReadProperties()
                                                         
       Dim Title As String
       Dim Subject As String
       Dim Comments As String
       Dim Author As String
                                                         
       Title = GetEXIFData(TargetFileName, EXIFImageTitle)
       Subject = GetEXIFData(TargetFileName, EXIFImageSubject)
       Comments = GetEXIFData(TargetFileName, EXIFImageComments)
       Author = GetEXIFData(TargetFileName, EXIFImageAuthor)
                                                         
       MsgBox "Title: " & Title & vbNewLine & _
              "Author: " & Author & vbNewLine & _
              "Subject: " & Subject & vbNewLine & _
              "Comments: " & Comments
   End Sub
                                                         
   Public Sub ImportEXIFToolsTable()
       
       Dim URL As String
       URL = "https://exiftool.org/TagNames/EXIF.html"
       Dim WB As Workbook
       Set WB = Application.Workbooks.Open(URL)
       
   End Sub
                                                         
   Public Function GetEXIFData(ByVal filename As String, ByVal PropertyName As PropertyNameEnum) As String
                                                         
       Dim Image               As Object
       Dim ImageProperty       As Object
       Dim Result              As String
                                                         
       Set Image = CreateObject("WIA.ImageFile")
       Image.LoadFile filename
                                                         
       For Each ImageProperty In Image.Properties
           If ImageProperty.PropertyID = PropertyName Then
               If TypeName(ImageProperty.value) = "String" Then
                   Result = ImageProperty.value
               Else
                   Result = Replace(StrConv(ImageProperty.value.BinaryData, vbUnicode), Chr(0), "")
               End If
               Exit For
           End If
       Next
                                                         
       GetEXIFData = Result
                                                         
       Set Image = Nothing
       Set ImageProperty = Nothing
                                                         
   End Function
                                                         
   Public Function WriteEXIFData(ByVal filename As String, ByVal PropertyName As PropertyNameEnum, ByVal PropertyValue As Variant, Optional ByVal OverWriteOriginal As Boolean = True, Optional ByVal CreateBackup As Boolean)
                                                         
       Dim Image               As Object
       Dim ImageProcess        As Object
       Dim ImageVector         As Object
       Dim NewFileName         As String
                                                         
       If CreateBackup = True Then
           Dim BackUpFilename  As String
           BackUpFilename = Replace(filename, ".jpg", "_BACKUP(" & Format(Now, "ddmmyyyy-hhnn") & ").jpg")
           FileCopy filename, BackUpFilename
       End If
                                                         
       Set Image = CreateObject("WIA.ImageFile")
       Set ImageProcess = CreateObject("WIA.ImageProcess")
       Set ImageVector = CreateObject("WIA.Vector")
                                                         
       Image.LoadFile filename
                                                         
       ImageProcess.Filters.Add ImageProcess.FilterInfos("Exif").FilterID
       ImageProcess.Filters(1).Properties("ID") = PropertyName
                                                         
       Select Case PropertyName
                                                         
           Case PropertyNameEnum.EXIFImageDateTimeOriginal
               Dim StringValue As String
               StringValue = Format(PropertyValue, "YYYY:MM:DD HH:MM:SS")
               ImageProcess.Filters(1).Properties("Type") = StringImagePropertyType
               ImageProcess.Filters(1).Properties("Value") = StringValue
                                                         
           Case Else
               ImageProcess.Filters(1).Properties("Type") = VectorOfBytesImagePropertyType
               ImageVector.SetFromString PropertyValue
               ImageProcess.Filters(1).Properties("Value") = ImageVector
                                                         
       End Select
                                                         
       Set Image = ImageProcess.Apply(Image)
                                                         
       If OverWriteOriginal = True Then
           NewFileName = filename
           Kill filename
       Else
           NewFileName = Replace(filename, ".jpg", "_metadata.jpg")
           If Len(Dir(NewFileName)) > 0 Then Kill NewFileName
       End If
                                                         
       Image.SaveFile NewFileName
                                                         
       WriteEXIFData = NewFileName
                                                         
       Set Image = Nothing
       Set ImageProcess = Nothing
       Set ImageVector = Nothing
                                                         
   End Function
                                                         