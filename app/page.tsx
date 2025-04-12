'use client';

import { useState, useRef, ChangeEvent, DragEvent, useEffect } from 'react';
import Image from 'next/image';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

type ExcelDataItem = {
	[key: string]: string;
};

export default function Home() {
	const [file, setFile] = useState<File | null>(null);
	const [isDragging, setIsDragging] = useState(false);
	const fileInputRef = useRef<HTMLInputElement>(null);
	const [excelData, setExcelData] = useState<ExcelDataItem[]>([]);
	const [headers, setHeaders] = useState<string[]>([]);
	const [isGenerating, setIsGenerating] = useState(false);
	const [backgroundImage, setBackgroundImage] =
		useState<HTMLImageElement | null>(null);
	const [isBackgroundLoaded, setIsBackgroundLoaded] = useState(false);
	const [backgroundError, setBackgroundError] = useState<string | null>(null);

	// Load background image when component mounts
	useEffect(() => {
		const img = new window.Image();
		img.src = '/assets/background.jpg';

		img.onload = () => {
			setBackgroundImage(img);
			setIsBackgroundLoaded(true);
			setBackgroundError(null);
		};

		img.onerror = () => {
			setBackgroundError(
				'Failed to load background image. Please ensure "assets/background.jpg" exists in the public folder.'
			);
			setIsBackgroundLoaded(false);
		};

		return () => {
			// Clean up
			img.onload = null;
			img.onerror = null;
		};
	}, []);

	const acceptedFileTypes = [
		'application/vnd.ms-excel',
		'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	];

	const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
		if (e.target.files && e.target.files[0]) {
			const selectedFile = e.target.files[0];
			if (isValidExcelFile(selectedFile)) {
				setFile(selectedFile);
				processExcelFile(selectedFile);
			} else {
				alert('Please upload only Excel files (.xls or .xlsx)');
			}
		}
	};

	const isValidExcelFile = (file: File): boolean => {
		return acceptedFileTypes.includes(file.type);
	};

	const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
		e.preventDefault();
		setIsDragging(true);
	};

	const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
		e.preventDefault();
		setIsDragging(false);
	};

	const handleDrop = (e: DragEvent<HTMLDivElement>) => {
		e.preventDefault();
		setIsDragging(false);

		if (e.dataTransfer.files && e.dataTransfer.files[0]) {
			const droppedFile = e.dataTransfer.files[0];
			if (isValidExcelFile(droppedFile)) {
				setFile(droppedFile);
				processExcelFile(droppedFile);
			} else {
				alert('Please upload only Excel files (.xls or .xlsx)');
			}
		}
	};

	const processExcelFile = async (file: File) => {
		const data = await file.arrayBuffer();
		const workbook = XLSX.read(data);
		const worksheet = workbook.Sheets[workbook.SheetNames[0]];

		// Convert to JSON
		const jsonData = XLSX.utils.sheet_to_json<ExcelDataItem>(worksheet, {
			raw: false,
		});

		// Extract headers from the first row
		if (jsonData.length > 0) {
			const firstRow = jsonData[0];
			setHeaders(Object.keys(firstRow));
			setExcelData(jsonData);
		}
	};

	const handleUploadClick = () => {
		fileInputRef.current?.click();
	};

	const generateImagesAndZip = async () => {
		if (!backgroundImage || excelData.length === 0) return;

		setIsGenerating(true);

		try {
			const zip = new JSZip();

			// For each row in excelData, create a new image
			for (let i = 0; i < excelData.length; i++) {
				const row = excelData[i];

				// Create a new canvas element
				const canvas = document.createElement('canvas');
				const ctx = canvas.getContext('2d');

				if (!ctx) continue;

				// Set canvas dimensions to match background image
				canvas.width = backgroundImage.width;
				canvas.height = backgroundImage.height;

				// Draw background image
				ctx.drawImage(backgroundImage, 0, 0);

				// Function to draw text with stroke (fixed position)
				const drawTextWithStroke = (
					text: string,
					x: number,
					y: number,
					fillStyle: string,
					strokeStyle: string,
					lineWidth: number,
					font: string
				) => {
					ctx.font = font;

					// Enhanced stroke settings for rounder appearance
					ctx.strokeStyle = strokeStyle;
					ctx.lineWidth = lineWidth;
					ctx.lineJoin = 'round'; // Makes corners rounded
					ctx.lineCap = 'round'; // Makes endpoints rounded
					ctx.miterLimit = 2; // Reduces pointiness at corners

					// Draw multiple strokes with slight offsets for softer edges
					for (let i = 0; i < 3; i++) {
						const offset = i * 0.5;

						// Drawing multiple overlapping strokes creates a softer appearance
						ctx.strokeText(text, x - offset, y);
						ctx.strokeText(text, x + offset, y);
						ctx.strokeText(text, x, y - offset);
						ctx.strokeText(text, x, y + offset);
					}

					// Draw fill on top
					ctx.fillStyle = fillStyle;
					ctx.fillText(text, x, y);
				};

				// Function to draw centered text with stroke
				const drawCenteredTextWithStroke = (
					text: string,
					centerX: number,
					y: number,
					fillStyle: string,
					strokeStyle: string,
					lineWidth: number,
					font: string
				) => {
					ctx.font = font;

					// Measure text width
					const textWidth = ctx.measureText(text).width;

					// Calculate x position to center the text
					const x = centerX - textWidth / 2;

					// Enhanced stroke settings for rounder appearance
					ctx.strokeStyle = strokeStyle;
					ctx.lineWidth = lineWidth;
					ctx.lineJoin = 'round'; // Makes corners rounded
					ctx.lineCap = 'round'; // Makes endpoints rounded
					ctx.miterLimit = 2; // Reduces pointiness at corners

					// Draw multiple strokes with slight offsets for softer edges
					for (let i = 0; i < 3; i++) {
						const offset = i * 0.5;

						// Drawing multiple overlapping strokes creates a softer appearance
						ctx.strokeText(text, x - offset, y);
						ctx.strokeText(text, x + offset, y);
						ctx.strokeText(text, x, y - offset);
						ctx.strokeText(text, x, y + offset);
					}

					// Draw fill on top
					ctx.fillStyle = fillStyle;
					ctx.fillText(text, x, y);
				};

				// Reference center point for the background image
				const centerX = backgroundImage.width / 2;

				// For broker field - BLACK with white stroke (FIXED POSITION)
				if (row['broker']) {
					drawTextWithStroke(
						row['broker'],
						530,
						500,
						'black', // fill color
						'white', // stroke color
						8, // stroke width
						"bold 130px 'Noto Sans TC', sans-serif"
					);
				}

				// For name field - YELLOW with black stroke (FIXED POSITION)
				if (row['name']) {
					drawTextWithStroke(
						row['name'],
						220,
						820,
						'#FFD700', // Yellow fill
						'black', // black stroke
						10, // stroke width
						"bold 240px 'Noto Sans TC', sans-serif"
					);
				}

				// For project name field - BLACK with white stroke (CENTERED)
				if (row['project name']) {
					drawCenteredTextWithStroke(
						row['project name'],
						centerX,
						1000,
						'black', // fill color
						'white', // stroke color
						5, // stroke width
						"bold 90px 'Noto Sans TC', sans-serif"
					);
				}

				// For submit field - RED with white stroke (CENTERED)
				if (row['submit']) {
					drawCenteredTextWithStroke(
						`FYP${row['submit']}`,
						centerX,
						1220,
						'#FF0000', // fill color
						'white', // stroke color
						8, // stroke width
						"bold 130px 'Noto Sans TC', sans-serif"
					);
				}

				// Add any additional fields
				headers.forEach((header, index) => {
					if (
						!['name', 'broker', 'project name', 'submit'].includes(
							header.toLowerCase()
						) &&
						row[header]
					) {
						drawTextWithStroke(
							`${header}: ${row[header]}`,
							100,
							300 + index * 50,
							'black', // fill color
							'white', // stroke color
							2, // stroke width
							"24px 'Noto Sans TC', sans-serif"
						);
					}
				});

				// Convert canvas to blob
				const blob = await new Promise<Blob>((resolve) => {
					canvas.toBlob(
						(blob) => {
							if (blob) resolve(blob);
							else resolve(new Blob([]));
						},
						'image/jpeg',
						0.95
					);
				});

				// Add to zip file
				zip.file(`image_${i + 1}.jpg`, blob);
			}

			// Generate and download zip
			const zipBlob = await zip.generateAsync({ type: 'blob' });
			saveAs(zipBlob, 'generated_images.zip');
		} catch (error) {
			console.error('Error generating images:', error);
			alert('Error generating images. Please try again.');
		} finally {
			setIsGenerating(false);
		}
	};

	const handleDownloadExampleFile = () => {
		// Create a link element
		const link = document.createElement('a');
		// Set the href to the example sheet in the public folder
		link.href = '/example-sheet.xlsx';
		// Set download attribute to specify the file name
		link.download = 'example-sheet.xlsx';
		// Add to document body
		document.body.appendChild(link);
		// Trigger the download
		link.click();
		// Clean up by removing the link
		document.body.removeChild(link);
	};

	return (
		<div
			className='flex flex-col items-center justify-center min-h-screen p-8'
			style={{ fontFamily: "'Noto Sans TC', sans-serif" }}
		>
			<div className='w-full max-w-4xl'>
				<h1 className='text-2xl font-bold mb-6 text-center'>
					Excel 檔案上傳
				</h1>

				{backgroundError && (
					<div className='bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4'>
						<p>{backgroundError}</p>
					</div>
				)}

				<div
					className={`border-2 border-dashed rounded-lg p-8 mb-4 transition-colors text-center
            ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300'}
            ${file ? 'bg-green-50' : ''}`}
					onDragOver={handleDragOver}
					onDragLeave={handleDragLeave}
					onDrop={handleDrop}
				>
					<div className='flex flex-col items-center'>
						{file ? (
							<>
								<div className='mb-4'>
									<Image
										src='/file.svg'
										width={48}
										height={48}
										alt='Excel file'
									/>
								</div>
								<p className='font-medium'>{file.name}</p>
								<p className='text-sm text-gray-500 mt-1'>
									{(file.size / 1024).toFixed(2)} KB
								</p>
							</>
						) : (
							<>
								<div className='mb-4'>
									<Image
										src='/file.svg'
										width={48}
										height={48}
										alt='Upload icon'
									/>
								</div>
								<p className='text-lg mb-2'>
									拖放 Excel 檔案到此處
								</p>
								<p className='text-sm text-gray-500'>
									僅支援 .xls 和 .xlsx 檔案格式
								</p>
							</>
						)}
					</div>
				</div>

				<input
					type='file'
					ref={fileInputRef}
					onChange={handleFileChange}
					accept='.xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
					className='hidden'
				/>

				<div className='flex flex-col sm:flex-row gap-4 mb-8'>
					<button
						onClick={handleDownloadExampleFile}
						className='py-3 px-4 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors flex-1'
					>
						下載範例檔案
					</button>

					<button
						onClick={handleUploadClick}
						className='py-3 px-4 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors flex-1'
					>
						{file ? '更換檔案' : '上傳 Excel 檔案'}
					</button>

					{excelData.length > 0 && (
						<button
							onClick={generateImagesAndZip}
							disabled={isGenerating || !isBackgroundLoaded}
							className={`py-3 px-4 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors flex-1 ${
								isGenerating || !isBackgroundLoaded
									? 'opacity-50 cursor-not-allowed'
									: ''
							}`}
						>
							{isGenerating
								? '正在產生...'
								: !isBackgroundLoaded
								? '正在載入背景...'
								: '產生圖片並下載'}
						</button>
					)}
				</div>

				{excelData.length > 0 && (
					<div className='overflow-x-auto'>
						<table className='min-w-full border-collapse border border-gray-300'>
							<thead>
								<tr className='bg-gray-100'>
									{headers.map((header, index) => (
										<th
											key={index}
											className='border border-gray-300 px-4 py-2 text-left text-black'
										>
											{header}
										</th>
									))}
								</tr>
							</thead>
							<tbody>
								{excelData.map((row, rowIndex) => (
									<tr
										key={rowIndex}
										className={
											rowIndex % 2 === 0
												? 'bg-white'
												: 'bg-gray-50'
										}
									>
										{headers.map((header, colIndex) => (
											<td
												key={colIndex}
												className='border border-gray-300 px-4 py-2 text-black'
											>
												{row[header]}
											</td>
										))}
									</tr>
								))}
							</tbody>
						</table>
					</div>
				)}
			</div>
		</div>
	);
}
