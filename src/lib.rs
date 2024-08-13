#![warn(clippy::all, rust_2018_idioms)]

mod app;
pub use app::App;

use std::collections::HashSet;

#[derive(Clone, Debug)]
struct Pairs {
    pairs: Vec<Vec<Option<usize>>>,
    size: usize,
}

impl Pairs {
    fn new(org_size: usize) -> Self {
        let size;
        if org_size % 2 != 0 {
            size = 1 + org_size;
        } else {
            size = org_size;
        }
        let mut pairs = vec![vec![None; size]; size];
        for i in 0..size {
            for j in 0..size {
                if i == j {
                    if i == size - 1 {
                        continue;
                    }
                    let nr = (i + j) % (size - 1);
                    pairs[i][size - 1] = Some(nr);
                    pairs[size - 1][j] = Some(nr);
                } else if let Some(_) = pairs[i][j] {
                    continue;
                } else {
                    pairs[i][j] = Some((i + j) % (size - 1));
                    pairs[j][i] = Some((i + j) % (size - 1));
                }
            }
        }
        Pairs { pairs, size }
    }
}

fn list_of_pairs(pairs: &Pairs, nr: usize) -> Vec<(usize, usize)> {
    let mut pair: Vec<(usize, usize)> = Vec::new();
    let mut taken: HashSet<usize> = HashSet::new();
    for i in 0..pairs.size {
        if taken.contains(&i) {
            continue;
        } else {
            for j in 0..pairs.size {
                if pairs.pairs[i][j] == Some(nr) {
                    pair.push((i, j));
                    taken.insert(j);
                }
            }
        }
    }
    pair
}

fn generate_pairs(mut names: Vec<&str>) -> Vec<Vec<(String, String)>> {
    let mut matrix_of_pairs = Vec::<Vec<(String, String)>>::new();
    if names.len() % 2 != 0 {
        names.push("");
    }
    let pair_matrix = Pairs::new(names.len());
    for nr in 0..pair_matrix.size - 1 {
        let mut pairs = Vec::<(String, String)>::new();
        let pairing = list_of_pairs(&pair_matrix, nr);
        for item in pairing.iter() {
            pairs.push((names[item.0].to_string(), names[item.1].to_string()));
        }
        matrix_of_pairs.push(pairs);
    }
    matrix_of_pairs
}

fn generate_list_of_names(namestr: &str) -> Vec<String> {
    //let list_of_names = Vec::new();
    let name_list =namestr
        .trim()
        .lines()
        .map(|line| {
            line.trim().to_string()
        })
        .collect::<Vec<String>>();
    let mut new_name_list = Vec::new();
    for name in name_list.iter() {
        new_name_list.push(name.replace(char::is_whitespace, " "));
    }
    new_name_list
}

use rust_xlsxwriter::*;

fn create_xlsx(all_pairs: Vec<Vec<(String,String)>>) -> Result<Vec<u8>, XlsxError> {
    
    let mut workbook = Workbook::new();

    for periode in 0..all_pairs.len() {
        let worksheet = workbook.add_worksheet().set_name(format!("{}",periode+1))?;
        let mut row: u32= 1;
        let mut grp_nr = 1;
        let mut col =1;
        for grp in all_pairs[periode].iter() {
            if (row/4) as f32  >= (all_pairs[0].len() / 2) as f32  {
                col += 4;
                row = 1;
            }
            worksheet.write_string(row, col, format!("Gruppe {}", grp_nr))?;
            worksheet.write_string(row+1, col, &grp.0)?;
            worksheet.write_string(row+2, col,&grp.1)?;
            grp_nr += 1;
            row += 4;
        } 
    }
 
    


    let buf = workbook.save_to_buffer()?;
    Ok(buf)
}